using Quantum;
using Quantum.Client.Windows;
using Quantum.Entities;
using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Dynamic;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Runtime.CompilerServices;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Web.Script.Serialization;
using System.Windows;
using System.Windows.Controls;
using Telerik.Windows.Controls;

class FetchFromJiraDialogContext : INotifyPropertyChanged
{
    string jiraServerUrl = "";
    string jiraEmail = "";
    string issueKey = "";
    bool saveCredentials = true;

    public event PropertyChangedEventHandler PropertyChanged;

    protected void OnPropertyChanged([CallerMemberName] string propertyName = null)
    {
        PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
    }

    protected bool SetProperty<T>(ref T field, T value, [CallerMemberName] string propertyName = null)
    {
        if (!EqualityComparer<T>.Default.Equals(field, value))
        {
            field = value;
            OnPropertyChanged(propertyName);
            return true;
        }
        return false;
    }

    public string JiraServerUrl
    {
        get => jiraServerUrl;
        set => SetProperty(ref jiraServerUrl, value);
    }

    public string JiraEmail
    {
        get => jiraEmail;
        set => SetProperty(ref jiraEmail, value);
    }

    public string IssueKey
    {
        get => issueKey;
        set => SetProperty(ref issueKey, value);
    }

    public bool SaveCredentials
    {
        get => saveCredentials;
        set => SetProperty(ref saveCredentials, value);
    }
}

class JiraIssue
{
    public string Key { get; set; }
    public string Summary { get; set; }
    public string Description { get; set; }
    public string IssueType { get; set; }
    public string Status { get; set; }
    public string Priority { get; set; }
}

async Task<JiraIssue> FetchJiraIssueAsync(string serverUrl, string email, string apiToken, string issueKey)
{
    using (var client = new HttpClient())
    {
        var credentials = Convert.ToBase64String(Encoding.ASCII.GetBytes($"{email}:{apiToken}"));
        client.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Basic", credentials);
        client.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));

        var url = $"{serverUrl.TrimEnd('/')}/rest/api/3/issue/{issueKey}";
        var response = await client.GetAsync(url);

        if (!response.IsSuccessStatusCode)
        {
            var errorContent = await response.Content.ReadAsStringAsync();
            throw new Exception($"Failed to fetch issue: {response.StatusCode}\n{errorContent}");
        }

        var jsonContent = await response.Content.ReadAsStringAsync();
        var serializer = new JavaScriptSerializer();
        var data = serializer.Deserialize<Dictionary<string, object>>(jsonContent);

        var fields = data["fields"] as Dictionary<string, object>;

        return new JiraIssue
        {
            Key = data["key"] as string,
            Summary = fields["summary"] as string,
            Description = ParseJiraDescription(fields.ContainsKey("description") ? fields["description"] : null),
            IssueType = GetNestedValue(fields, "issuetype", "name") as string,
            Status = GetNestedValue(fields, "status", "name") as string,
            Priority = fields.ContainsKey("priority") && fields["priority"] != null
                ? GetNestedValue(fields, "priority", "name") as string
                : "None"
        };
    }
}

object GetNestedValue(Dictionary<string, object> dict, params string[] keys)
{
    object current = dict;
    foreach (var key in keys)
    {
        if (current is Dictionary<string, object> currentDict && currentDict.ContainsKey(key))
        {
            current = currentDict[key];
        }
        else
        {
            return null;
        }
    }
    return current;
}

string ParseJiraDescription(object descriptionObj)
{
    if (descriptionObj == null)
        return "";

    if (descriptionObj is string)
        return descriptionObj as string;

    if (descriptionObj is Dictionary<string, object> descDict)
    {
        if (descDict.ContainsKey("content") && descDict["content"] is ArrayList contentArray)
        {
            var sb = new StringBuilder();
            foreach (var item in contentArray)
            {
                if (item is Dictionary<string, object> itemDict && itemDict.ContainsKey("content"))
                {
                    if (itemDict["content"] is ArrayList innerContentArray)
                    {
                        foreach (var textItem in innerContentArray)
                        {
                            if (textItem is Dictionary<string, object> textDict && textDict.ContainsKey("text"))
                            {
                                sb.AppendLine(textDict["text"] as string);
                            }
                        }
                    }
                }
            }
            return sb.ToString().Trim();
        }
    }

    return descriptionObj.ToString();
}

async Task CreateWorkItemFromJiraIssueAsync(JiraIssue jiraIssue)
{
    if (jiraTransactor == null)
        throw new InvalidOperationException("Transactor not initialized. Extension may not be fully loaded.");

    var workItemName = $"[{jiraIssue.Key}] {jiraIssue.Summary}";
    var workItemNotes = $"Jira Issue: {jiraIssue.Key}\n" +
                       $"Type: {jiraIssue.IssueType}\n" +
                       $"Status: {jiraIssue.Status}\n" +
                       $"Priority: {jiraIssue.Priority}\n\n" +
                       $"Description:\n{jiraIssue.Description}";

    var currentSnapshot = await jiraTransactor.SnapshotAsync();

    var attributes = new Dictionary<Guid, Quantum.Entities.Attribute>();
    foreach (var kvp in currentSnapshot.Attributes)
        attributes.Add(kvp.Key, kvp.Value);

    var attributeValues = new Dictionary<AttributeObjectCompositeKey, object>();
    foreach (var kvp in currentSnapshot.AttributeValues)
        attributeValues.Add(kvp.Key, kvp.Value);

    var enumerationValues = new Dictionary<Guid, AttributeCorrelation<EnumerationValue>>();
    foreach (var kvp in currentSnapshot.EnumerationValues)
        enumerationValues.Add(kvp.Key, kvp.Value);

    var interests = new Dictionary<Guid, ItemPersonCorrelation<Interest>>();
    foreach (var kvp in currentSnapshot.Interests)
        interests.Add(kvp.Key, kvp.Value);

    var items = new Dictionary<Guid, Item>();
    foreach (var kvp in currentSnapshot.Items)
        items.Add(kvp.Key, kvp.Value);
    items.Add(Guid.NewGuid(), new Item(workItemName, workItemNotes));

    var people = new Dictionary<Guid, Person>();
    foreach (var kvp in currentSnapshot.People)
        people.Add(kvp.Key, kvp.Value);

    var periods = new Dictionary<Guid, ItemPersonCorrelation<Period>>();
    foreach (var kvp in currentSnapshot.Periods)
        periods.Add(kvp.Key, kvp.Value);

    var periodsTiming = new List<Guid>();
    foreach (var item in currentSnapshot.PeriodsTiming)
        periodsTiming.Add(item);

    var targetFrame = new Quantum.Entities.Frame(
        attributes,
        attributeValues,
        enumerationValues,
        interests,
        items,
        people,
        periods,
        periodsTiming
    );

    var changeSet = currentSnapshot.Difference(targetFrame);
    await jiraTransactor.ApplyChangesAsync(changeSet);
}

class JiraTransactor : Transactor
{
}

JiraTransactor jiraTransactor;

void EngineStartingHandler(object sender, EngineStartingEventArgs e)
{
    jiraTransactor = new JiraTransactor();
    e.AddTransactor(jiraTransactor);
}

async void FetchFromJiraClick(object sender, RoutedEventArgs e)
{
    await Extension.OnUiThreadAsync(async () =>
    {
        var dialog = (Window)Extension.LoadUiElement("FetchFromJiraDialog.xaml");
        var context = new FetchFromJiraDialogContext();

        var savedServerUrl = Extension.DatabaseStorage?.Get<string>("JiraServerUrl");
        var savedEmail = Extension.DatabaseStorage?.Get<string>("JiraEmail");

        if (!string.IsNullOrEmpty(savedServerUrl))
            context.JiraServerUrl = savedServerUrl;
        if (!string.IsNullOrEmpty(savedEmail))
            context.JiraEmail = savedEmail;

        dialog.DataContext = context;

        var fetchButton = (Button)dialog.FindName("fetchButton");
        var cancelButton = (Button)dialog.FindName("cancelButton");
        var apiTokenBox = (PasswordBox)dialog.FindName("jiraApiToken");
        var statusMessage = (TextBlock)dialog.FindName("statusMessage");

        var savedApiToken = Extension.DatabaseStorage?.Get<string>("JiraApiToken");
        if (!string.IsNullOrEmpty(savedApiToken))
            apiTokenBox.Password = savedApiToken;

        fetchButton.Click += async (s, args) =>
        {
            try
            {
                fetchButton.IsEnabled = false;
                statusMessage.Text = "Fetching issue from Jira...";

                var serverUrl = context.JiraServerUrl?.Trim();
                var email = context.JiraEmail?.Trim();
                var apiToken = apiTokenBox.Password?.Trim();
                var issueKey = context.IssueKey?.Trim();

                if (string.IsNullOrEmpty(serverUrl) || string.IsNullOrEmpty(email) ||
                    string.IsNullOrEmpty(apiToken) || string.IsNullOrEmpty(issueKey))
                {
                    MessageDialog.Present(dialog, "Please fill in all fields.", "Missing Information", MessageBoxImage.Warning);
                    statusMessage.Text = "";
                    fetchButton.IsEnabled = true;
                    return;
                }

                var jiraIssue = await FetchJiraIssueAsync(serverUrl, email, apiToken, issueKey);

                statusMessage.Text = "Creating work item in Grindstone...";
                await CreateWorkItemFromJiraIssueAsync(jiraIssue);

                if (context.SaveCredentials && Extension.DatabaseStorage != null)
                {
                    Extension.DatabaseStorage.Set("JiraServerUrl", serverUrl);
                    Extension.DatabaseStorage.Set("JiraEmail", email);
                    Extension.DatabaseStorage.Set("JiraApiToken", apiToken);
                }

                MessageDialog.Present(dialog,
                    $"Successfully created work item for {jiraIssue.Key}:\n{jiraIssue.Summary}",
                    "Success",
                    MessageBoxImage.Information);

                dialog.DialogResult = true;
            }
            catch (Exception ex)
            {
                statusMessage.Text = "";
                MessageDialog.Present(dialog, $"Error: {ex.Message}", "Error Fetching Issue", MessageBoxImage.Error);
                fetchButton.IsEnabled = true;
            }
        };

        cancelButton.Click += (s, args) => dialog.DialogResult = false;

        dialog.ShowDialog();
    });
}

var extensionsMenuExtensionId = Guid.Parse("{27F65593-7235-4108-B5D9-F0DE417D8536}");

Extension.EngineStarting += EngineStartingHandler;

await Extension.OnUiThreadAsync(() =>
{
    var jiraMenuItem = new RadMenuItem { Header = "Jira Integration" };

    var fetchMenuItem = new RadMenuItem { Header = "Fetch from Jira" };
    fetchMenuItem.Click += FetchFromJiraClick;

    jiraMenuItem.Items.Add(fetchMenuItem);

    Extension.PostMessage(extensionsMenuExtensionId, jiraMenuItem);
});
