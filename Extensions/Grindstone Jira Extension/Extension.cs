using Quantum;
using Quantum.Client.Windows;
using Quantum.Entities;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Linq;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Runtime.CompilerServices;
using System.Text;
using System.Threading.Tasks;
using System.Web.Script.Serialization;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using Telerik.Windows.Controls;

class JiraConnection
{
    public string Id { get; set; }
    public string Name { get; set; }
    public string ServerUrl { get; set; }
    public string Email { get; set; }
    public string ApiToken { get; set; }

    public JiraConnection()
    {
        Id = Guid.NewGuid().ToString();
    }
}

const string JiraConnectionsStorageKey = "JiraConnections";

List<JiraConnection> LoadJiraConnections()
{
    var serializer = new JavaScriptSerializer();
    var json = Extension.DatabaseStorage?.Get<string>(JiraConnectionsStorageKey);
    if (string.IsNullOrEmpty(json))
        return new List<JiraConnection>();

    try
    {
        var data = serializer.Deserialize<List<Dictionary<string, object>>>(json);
        return data.Select(d => new JiraConnection
        {
            Id = d["Id"] as string,
            Name = d["Name"] as string,
            ServerUrl = d["ServerUrl"] as string,
            Email = d["Email"] as string,
            ApiToken = d["ApiToken"] as string
        }).ToList();
    }
    catch
    {
        return new List<JiraConnection>();
    }
}

void SaveJiraConnections(List<JiraConnection> connections)
{
    var serializer = new JavaScriptSerializer();
    var data = connections.Select(c => new Dictionary<string, object>
    {
        { "Id", c.Id },
        { "Name", c.Name },
        { "ServerUrl", c.ServerUrl },
        { "Email", c.Email },
        { "ApiToken", c.ApiToken }
    }).ToList();
    var json = serializer.Serialize(data);
    Extension.DatabaseStorage?.Set(JiraConnectionsStorageKey, json);
}

class FetchFromJiraDialogContext : INotifyPropertyChanged
{
    string issueKey = "";

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

    public string IssueKey
    {
        get => issueKey;
        set => SetProperty(ref issueKey, value);
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

async Task<bool> TestJiraConnectionAsync(string serverUrl, string email, string apiToken)
{
    using (var client = new HttpClient())
    {
        var credentials = Convert.ToBase64String(Encoding.ASCII.GetBytes($"{email}:{apiToken}"));
        client.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Basic", credentials);
        client.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));

        var url = $"{serverUrl.TrimEnd('/')}/rest/api/3/myself";
        var response = await client.GetAsync(url);

        return response.IsSuccessStatusCode;
    }
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

async void ManageConnectionsClick(object sender, RoutedEventArgs e)
{
    await Extension.OnUiThreadAsync(() =>
    {
        var dialog = (Window)Extension.LoadUiElement("ManageConnectionsDialog.xaml");
        var connections = LoadJiraConnections();

        var connectionsList = (ListBox)dialog.FindName("connectionsList");
        var connectionName = (TextBox)dialog.FindName("connectionName");
        var serverUrl = (TextBox)dialog.FindName("serverUrl");
        var email = (TextBox)dialog.FindName("email");
        var apiToken = (PasswordBox)dialog.FindName("apiToken");
        var addButton = (Button)dialog.FindName("addButton");
        var updateButton = (Button)dialog.FindName("updateButton");
        var deleteButton = (Button)dialog.FindName("deleteButton");
        var testButton = (Button)dialog.FindName("testButton");
        var closeButton = (Button)dialog.FindName("closeButton");

        connectionsList.ItemsSource = connections;

        connectionsList.SelectionChanged += (s, args) =>
        {
            var selectedConnection = connectionsList.SelectedItem as JiraConnection;
            if (selectedConnection != null)
            {
                connectionName.Text = selectedConnection.Name;
                serverUrl.Text = selectedConnection.ServerUrl;
                email.Text = selectedConnection.Email;
                apiToken.Password = selectedConnection.ApiToken;
                updateButton.IsEnabled = true;
                deleteButton.IsEnabled = true;
                testButton.IsEnabled = true;
            }
            else
            {
                updateButton.IsEnabled = false;
                deleteButton.IsEnabled = false;
                testButton.IsEnabled = false;
            }
        };

        addButton.Click += (s, args) =>
        {
            if (string.IsNullOrWhiteSpace(connectionName.Text) ||
                string.IsNullOrWhiteSpace(serverUrl.Text) ||
                string.IsNullOrWhiteSpace(email.Text) ||
                string.IsNullOrWhiteSpace(apiToken.Password))
            {
                MessageDialog.Present(dialog, "Please fill in all fields.", "Missing Information", MessageBoxImage.Warning);
                return;
            }

            var newConnection = new JiraConnection
            {
                Name = connectionName.Text.Trim(),
                ServerUrl = serverUrl.Text.Trim(),
                Email = email.Text.Trim(),
                ApiToken = apiToken.Password.Trim()
            };

            connections.Add(newConnection);
            SaveJiraConnections(connections);

            connectionsList.ItemsSource = null;
            connectionsList.ItemsSource = connections;

            connectionName.Clear();
            serverUrl.Clear();
            email.Clear();
            apiToken.Clear();
        };

        updateButton.Click += (s, args) =>
        {
            var selectedConnection = connectionsList.SelectedItem as JiraConnection;
            if (selectedConnection == null)
                return;

            if (string.IsNullOrWhiteSpace(connectionName.Text) ||
                string.IsNullOrWhiteSpace(serverUrl.Text) ||
                string.IsNullOrWhiteSpace(email.Text) ||
                string.IsNullOrWhiteSpace(apiToken.Password))
            {
                MessageDialog.Present(dialog, "Please fill in all fields.", "Missing Information", MessageBoxImage.Warning);
                return;
            }

            selectedConnection.Name = connectionName.Text.Trim();
            selectedConnection.ServerUrl = serverUrl.Text.Trim();
            selectedConnection.Email = email.Text.Trim();
            selectedConnection.ApiToken = apiToken.Password.Trim();

            SaveJiraConnections(connections);

            connectionsList.ItemsSource = null;
            connectionsList.ItemsSource = connections;
        };

        deleteButton.Click += (s, args) =>
        {
            var selectedConnection = connectionsList.SelectedItem as JiraConnection;
            if (selectedConnection == null)
                return;

            var result = MessageBox.Show(
                $"Are you sure you want to delete the connection '{selectedConnection.Name}'?",
                "Confirm Delete",
                MessageBoxButton.YesNo,
                MessageBoxImage.Question);

            if (result == MessageBoxResult.Yes)
            {
                connections.Remove(selectedConnection);
                SaveJiraConnections(connections);

                connectionsList.ItemsSource = null;
                connectionsList.ItemsSource = connections;

                connectionName.Clear();
                serverUrl.Clear();
                email.Clear();
                apiToken.Clear();
            }
        };

        testButton.Click += async (s, args) =>
        {
            if (string.IsNullOrWhiteSpace(serverUrl.Text) ||
                string.IsNullOrWhiteSpace(email.Text) ||
                string.IsNullOrWhiteSpace(apiToken.Password))
            {
                MessageDialog.Present(dialog, "Please fill in Server URL, Email, and API Token to test the connection.", "Missing Information", MessageBoxImage.Warning);
                return;
            }

            testButton.IsEnabled = false;
            testButton.Content = "Testing...";

            try
            {
                var success = await TestJiraConnectionAsync(
                    serverUrl.Text.Trim(),
                    email.Text.Trim(),
                    apiToken.Password.Trim());

                if (success)
                {
                    MessageDialog.Present(dialog, "Connection successful! You can authenticate with this Jira server.", "Test Successful", MessageBoxImage.Information);
                }
                else
                {
                    MessageDialog.Present(dialog, "Connection failed. Please check your Server URL, Email, and API Token.", "Test Failed", MessageBoxImage.Error);
                }
            }
            catch (Exception ex)
            {
                MessageDialog.Present(dialog, $"Connection test failed:\n{ex.Message}", "Test Error", MessageBoxImage.Error);
            }
            finally
            {
                testButton.IsEnabled = true;
                testButton.Content = "Test";
            }
        };

        closeButton.Click += (s, args) => dialog.Close();

        updateButton.IsEnabled = false;
        deleteButton.IsEnabled = false;
        testButton.IsEnabled = false;

        dialog.ShowDialog();
    });
}

async void FetchFromJiraClick(object sender, RoutedEventArgs e)
{
    await Extension.OnUiThreadAsync(async () =>
    {
        var connections = LoadJiraConnections();

        if (connections.Count == 0)
        {
            MessageDialog.Present(
                "No Jira connections configured. Please use 'Manage Connections' to add a Jira server first.",
                "No Connections",
                MessageBoxImage.Warning);
            return;
        }

        var dialog = (Window)Extension.LoadUiElement("FetchFromJiraDialog.xaml");
        var context = new FetchFromJiraDialogContext();
        dialog.DataContext = context;

        var connectionComboBox = (ComboBox)dialog.FindName("connectionComboBox");
        var fetchButton = (Button)dialog.FindName("fetchButton");
        var cancelButton = (Button)dialog.FindName("cancelButton");
        var statusMessage = (TextBlock)dialog.FindName("statusMessage");

        connectionComboBox.ItemsSource = connections;
        if (connections.Count > 0)
            connectionComboBox.SelectedIndex = 0;

        fetchButton.Click += async (s, args) =>
        {
            try
            {
                fetchButton.IsEnabled = false;
                statusMessage.Text = "Fetching issue from Jira...";

                var selectedConnection = connectionComboBox.SelectedItem as JiraConnection;
                var issueKey = context.IssueKey?.Trim();

                if (selectedConnection == null)
                {
                    MessageDialog.Present(dialog, "Please select a Jira connection.", "Missing Connection", MessageBoxImage.Warning);
                    statusMessage.Text = "";
                    fetchButton.IsEnabled = true;
                    return;
                }

                if (string.IsNullOrEmpty(issueKey))
                {
                    MessageDialog.Present(dialog, "Please enter an issue key.", "Missing Issue Key", MessageBoxImage.Warning);
                    statusMessage.Text = "";
                    fetchButton.IsEnabled = true;
                    return;
                }

                var jiraIssue = await FetchJiraIssueAsync(
                    selectedConnection.ServerUrl,
                    selectedConnection.Email,
                    selectedConnection.ApiToken,
                    issueKey);

                statusMessage.Text = "Creating work item in Grindstone...";
                await CreateWorkItemFromJiraIssueAsync(jiraIssue);

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

    var manageConnectionsMenuItem = new RadMenuItem { Header = "Manage Connections" };
    manageConnectionsMenuItem.Click += ManageConnectionsClick;

    var fetchMenuItem = new RadMenuItem { Header = "Fetch from Jira" };
    fetchMenuItem.Click += FetchFromJiraClick;

    jiraMenuItem.Items.Add(manageConnectionsMenuItem);
    jiraMenuItem.Items.Add(fetchMenuItem);

    Extension.PostMessage(extensionsMenuExtensionId, jiraMenuItem);
});
