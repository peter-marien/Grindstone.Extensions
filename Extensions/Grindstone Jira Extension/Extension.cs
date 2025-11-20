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

void ShowErrorDialog(string title, string message, Exception ex)
{
    Extension.OnUiThreadAsync(() =>
    {
        var dialog = (Window)Extension.LoadUiElement("ErrorDialog.xaml");

        var errorTitle = (TextBlock)dialog.FindName("errorTitle");
        var errorMessage = (TextBlock)dialog.FindName("errorMessage");
        var errorDetails = (TextBox)dialog.FindName("errorDetails");
        var copyButton = (Button)dialog.FindName("copyButton");
        var closeButton = (Button)dialog.FindName("closeButton");

        errorTitle.Text = title;
        errorMessage.Text = message;
        errorDetails.Text = ex.ToString();

        copyButton.Click += (s, e) =>
        {
            try
            {
                var fullError = $"{title}\n\n{message}\n\nDetails:\n{ex.ToString()}";
                Clipboard.SetText(fullError);
                copyButton.Content = "Copied!";
                Task.Delay(2000).ContinueWith(_ =>
                {
                    Extension.OnUiThreadAsync(() => copyButton.Content = "Copy Error");
                });
            }
            catch
            {
            }
        };

        closeButton.Click += (s, e) => dialog.Close();

        dialog.ShowDialog();
    });
}

List<JiraConnection> LoadJiraConnections()
{
    var serializer = new JavaScriptSerializer { MaxJsonLength = int.MaxValue };
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
    var serializer = new JavaScriptSerializer { MaxJsonLength = int.MaxValue };
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

class CreateWorkItemFromJiraDialogContext : INotifyPropertyChanged
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

class JiraWorklog
{
    public string IssueKey { get; set; }
    public string IssueSummary { get; set; }
    public string Comment { get; set; }
    public DateTime Started { get; set; }
    public int TimeSpentSeconds { get; set; }
}

class WorklogDisplayItem
{
    public string StartTime { get; set; }
    public string EndTime { get; set; }
    public string Duration { get; set; }
    public string WorkItemName { get; set; }
    public string Notes { get; set; }
    public Guid PeriodId { get; set; }
    public Guid ItemId { get; set; }
    public string JiraKey { get; set; }
    public string JiraConnection { get; set; }
}

async Task<bool> SyncWorklogToJiraAsync(string serverUrl, string email, string apiToken, string issueKey, DateTime started, int timeSpentSeconds, string comment)
{
    using (var client = new HttpClient())
    {
        var credentials = Convert.ToBase64String(Encoding.ASCII.GetBytes($"{email}:{apiToken}"));
        client.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Basic", credentials);
        client.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));

        var url = $"{serverUrl.TrimEnd('/')}/rest/api/3/issue/{issueKey}/worklog";

        var worklogData = new Dictionary<string, object>
        {
            { "started", started.ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ss.fff+0000") },
            { "timeSpentSeconds", timeSpentSeconds },
            { "comment", new Dictionary<string, object>
                {
                    { "type", "doc" },
                    { "version", 1 },
                    { "content", new object[]
                        {
                            new Dictionary<string, object>
                            {
                                { "type", "paragraph" },
                                { "content", new object[]
                                    {
                                        new Dictionary<string, object>
                                        {
                                            { "type", "text" },
                                            { "text", comment ?? "" }
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
            }
        };

        var serializer = new JavaScriptSerializer { MaxJsonLength = int.MaxValue };
        var jsonContent = serializer.Serialize(worklogData);
        var content = new StringContent(jsonContent, Encoding.UTF8, "application/json");

        var response = await client.PostAsync(url, content);

        if (!response.IsSuccessStatusCode)
        {
            var errorContent = await response.Content.ReadAsStringAsync();
            throw new Exception($"Failed to sync worklog: {response.StatusCode}\n{errorContent}");
        }

        return true;
    }
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

async Task<List<JiraWorklog>> ImportWorklogsFromJiraAsync(string serverUrl, string email, string apiToken, DateTime date)
{
    var worklogs = new List<JiraWorklog>();

    using (var client = new HttpClient())
    {
        var credentials = Convert.ToBase64String(Encoding.ASCII.GetBytes($"{email}:{apiToken}"));
        client.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Basic", credentials);
        client.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));

        var startTimestamp = new DateTimeOffset(date.Date).ToUnixTimeMilliseconds();
        var endTimestamp = new DateTimeOffset(date.Date.AddDays(1)).ToUnixTimeMilliseconds();

        var url = $"{serverUrl.TrimEnd('/')}/rest/api/3/worklog/updated?since={startTimestamp}";

        var response = await client.GetAsync(url);

        if (!response.IsSuccessStatusCode)
        {
            var errorContent = await response.Content.ReadAsStringAsync();
            throw new Exception($"Failed to fetch worklogs: {response.StatusCode}\n{errorContent}");
        }

        var serializer = new JavaScriptSerializer { MaxJsonLength = int.MaxValue };
        var jsonContent = await response.Content.ReadAsStringAsync();
        var data = serializer.Deserialize<Dictionary<string, object>>(jsonContent);

        var userEmail = email.ToLower();

        if (data.ContainsKey("values") && data["values"] is ArrayList values)
        {
            foreach (var valueObj in values)
            {
                if (valueObj is Dictionary<string, object> value)
                {
                    var worklogId = value.ContainsKey("worklogId") ? value["worklogId"].ToString() : null;
                    var properties = value.ContainsKey("properties") ? value["properties"] as ArrayList : null;

                    if (properties != null)
                    {
                        foreach (var propObj in properties)
                        {
                            if (propObj is Dictionary<string, object> prop)
                            {
                                var issueKey = prop.ContainsKey("key") ? prop["key"] as string : null;

                                if (!string.IsNullOrEmpty(issueKey) && !string.IsNullOrEmpty(worklogId))
                                {
                                    var worklogDetails = await FetchSingleWorklogAsync(client, serverUrl, issueKey, worklogId, date, userEmail, serializer);
                                    if (worklogDetails != null)
                                    {
                                        worklogs.Add(worklogDetails);
                                    }
                                }
                            }
                        }
                    }
                }
            }
        }

        var jql = $"worklogDate = '{date:yyyy-MM-dd}' AND worklogAuthor = currentUser()";
        var searchUrl = $"{serverUrl.TrimEnd('/')}/rest/api/3/search/jql";

        var requestBody = new Dictionary<string, object>
        {
            { "jql", jql },
            { "fields", new[] { "summary" } },
            { "maxResults", 1000 }
        };

        var jsonRequestBody = serializer.Serialize(requestBody);
        var content = new StringContent(jsonRequestBody, Encoding.UTF8, "application/json");

        var searchResponse = await client.PostAsync(searchUrl, content);

        if (searchResponse.IsSuccessStatusCode)
        {
            var searchContent = await searchResponse.Content.ReadAsStringAsync();
            var searchData = serializer.Deserialize<Dictionary<string, object>>(searchContent);

            if (searchData.ContainsKey("issues") && searchData["issues"] is ArrayList issues)
            {
                foreach (var issueObj in issues)
                {
                    if (issueObj is Dictionary<string, object> issue)
                    {
                        var issueKey = issue["key"] as string;
                        var fields = issue["fields"] as Dictionary<string, object>;
                        var issueSummary = fields != null && fields.ContainsKey("summary") ? fields["summary"] as string : "";

                        var issueWorklogsUrl = $"{serverUrl.TrimEnd('/')}/rest/api/3/issue/{issueKey}/worklog";
                        var worklogResponse = await client.GetAsync(issueWorklogsUrl);

                        if (worklogResponse.IsSuccessStatusCode)
                        {
                            var worklogContent = await worklogResponse.Content.ReadAsStringAsync();
                            var worklogData = serializer.Deserialize<Dictionary<string, object>>(worklogContent);

                            if (worklogData.ContainsKey("worklogs") && worklogData["worklogs"] is ArrayList worklogArray)
                            {
                                foreach (var worklogObj in worklogArray)
                                {
                                    if (worklogObj is Dictionary<string, object> worklog)
                                    {
                                        var author = worklog.ContainsKey("author") ? worklog["author"] as Dictionary<string, object> : null;
                                        var authorEmail = author != null && author.ContainsKey("emailAddress") ? (author["emailAddress"] as string)?.ToLower() : "";

                                        if (authorEmail == userEmail)
                                        {
                                            var started = worklog.ContainsKey("started") ? worklog["started"] as string : null;
                                            var timeSpentSeconds = worklog.ContainsKey("timeSpentSeconds") ? Convert.ToInt32(worklog["timeSpentSeconds"]) : 0;
                                            var comment = "";

                                            if (worklog.ContainsKey("comment"))
                                            {
                                                comment = ParseJiraDescription(worklog["comment"]);
                                            }

                                            if (!string.IsNullOrEmpty(started))
                                            {
                                                var startedDate = DateTime.Parse(started).ToUniversalTime();
                                                if (startedDate.Date == date.Date)
                                                {
                                                    var worklogId = worklog.ContainsKey("id") ? worklog["id"].ToString() : "";
                                                    var existingWorklog = worklogs.FirstOrDefault(w =>
                                                        w.IssueKey == issueKey &&
                                                        w.Started == startedDate &&
                                                        w.TimeSpentSeconds == timeSpentSeconds);

                                                    if (existingWorklog == null)
                                                    {
                                                        worklogs.Add(new JiraWorklog
                                                        {
                                                            IssueKey = issueKey,
                                                            IssueSummary = issueSummary,
                                                            Comment = comment,
                                                            Started = startedDate,
                                                            TimeSpentSeconds = timeSpentSeconds
                                                        });
                                                    }
                                                }
                                            }
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
            }
        }
    }

    return worklogs;
}

async Task<JiraWorklog> FetchSingleWorklogAsync(HttpClient client, string serverUrl, string issueKey, string worklogId, DateTime date, string userEmail, JavaScriptSerializer serializer)
{
    try
    {
        var url = $"{serverUrl.TrimEnd('/')}/rest/api/3/issue/{issueKey}/worklog/{worklogId}";
        var response = await client.GetAsync(url);

        if (response.IsSuccessStatusCode)
        {
            var content = await response.Content.ReadAsStringAsync();
            var worklog = serializer.Deserialize<Dictionary<string, object>>(content);

            var author = worklog.ContainsKey("author") ? worklog["author"] as Dictionary<string, object> : null;
            var authorEmail = author != null && author.ContainsKey("emailAddress") ? (author["emailAddress"] as string)?.ToLower() : "";

            if (authorEmail == userEmail)
            {
                var started = worklog.ContainsKey("started") ? worklog["started"] as string : null;
                var timeSpentSeconds = worklog.ContainsKey("timeSpentSeconds") ? Convert.ToInt32(worklog["timeSpentSeconds"]) : 0;
                var comment = "";

                if (worklog.ContainsKey("comment"))
                {
                    comment = ParseJiraDescription(worklog["comment"]);
                }

                if (!string.IsNullOrEmpty(started))
                {
                    var startedDate = DateTime.Parse(started).ToUniversalTime();
                    if (startedDate.Date == date.Date)
                    {
                        var issueUrl = $"{serverUrl.TrimEnd('/')}/rest/api/3/issue/{issueKey}?fields=summary";
                        var issueResponse = await client.GetAsync(issueUrl);
                        var issueSummary = "";

                        if (issueResponse.IsSuccessStatusCode)
                        {
                            var issueContent = await issueResponse.Content.ReadAsStringAsync();
                            var issue = serializer.Deserialize<Dictionary<string, object>>(issueContent);
                            var fields = issue.ContainsKey("fields") ? issue["fields"] as Dictionary<string, object> : null;
                            issueSummary = fields != null && fields.ContainsKey("summary") ? fields["summary"] as string : "";
                        }

                        return new JiraWorklog
                        {
                            IssueKey = issueKey,
                            IssueSummary = issueSummary,
                            Comment = comment,
                            Started = startedDate,
                            TimeSpentSeconds = timeSpentSeconds
                        };
                    }
                }
            }
        }
    }
    catch
    {
    }

    return null;
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
        var serializer = new JavaScriptSerializer { MaxJsonLength = int.MaxValue };
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

async Task<Guid> FindOrCreateWorkItemForJiraKeyAsync(string jiraKey, string issueSummary, string connectionName)
{
    if (jiraTransactor == null)
        throw new InvalidOperationException("Transactor not initialized. Extension may not be fully loaded.");

    var currentSnapshot = await jiraTransactor.SnapshotAsync();

    var jiraKeyAttributeId = Guid.Empty;
    foreach (var attr in currentSnapshot.Attributes)
    {
        if (attr.Value.Name == "Jira Key")
        {
            jiraKeyAttributeId = attr.Key;
            break;
        }
    }

    if (jiraKeyAttributeId != Guid.Empty)
    {
        foreach (var attrValue in currentSnapshot.AttributeValues)
        {
            if (attrValue.Key.AttributeId == jiraKeyAttributeId && attrValue.Value as string == jiraKey)
            {
                return attrValue.Key.ObjectId;
            }
        }
    }

    var jiraIssue = new JiraIssue
    {
        Key = jiraKey,
        Summary = issueSummary,
        Description = "",
        IssueType = "",
        Status = "",
        Priority = ""
    };

    return await CreateWorkItemFromJiraIssueAsyncInternal(jiraIssue, connectionName);
}

async Task<Guid> CreateWorkItemFromJiraIssueAsyncInternal(JiraIssue jiraIssue, string connectionName)
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

    var jiraKeyAttributeId = Guid.Empty;
    var jiraConnectionAttributeId = Guid.Empty;

    foreach (var attr in attributes)
    {
        if (attr.Value.Name == "Jira Key")
            jiraKeyAttributeId = attr.Key;
        if (attr.Value.Name == "Jira Connection")
            jiraConnectionAttributeId = attr.Key;
    }

    if (jiraKeyAttributeId == Guid.Empty)
    {
        jiraKeyAttributeId = Guid.NewGuid();
        attributes.Add(jiraKeyAttributeId, new Quantum.Entities.Attribute("Jira Key", "The Jira issue key for this work item"));
    }

    if (jiraConnectionAttributeId == Guid.Empty)
    {
        jiraConnectionAttributeId = Guid.NewGuid();
        attributes.Add(jiraConnectionAttributeId, new Quantum.Entities.Attribute("Jira Connection", "The Jira connection used to fetch this work item"));
    }

    var attributeValues = new Dictionary<AttributeObjectCompositeKey, object>();
    foreach (var kvp in currentSnapshot.AttributeValues)
        attributeValues.Add(kvp.Key, kvp.Value);

    var newWorkItemId = Guid.NewGuid();

    attributeValues.Add(new AttributeObjectCompositeKey(jiraKeyAttributeId, newWorkItemId), jiraIssue.Key);
    attributeValues.Add(new AttributeObjectCompositeKey(jiraConnectionAttributeId, newWorkItemId), connectionName);

    var enumerationValues = new Dictionary<Guid, AttributeCorrelation<EnumerationValue>>();
    foreach (var kvp in currentSnapshot.EnumerationValues)
        enumerationValues.Add(kvp.Key, kvp.Value);

    var interests = new Dictionary<Guid, ItemPersonCorrelation<Interest>>();
    foreach (var kvp in currentSnapshot.Interests)
        interests.Add(kvp.Key, kvp.Value);

    var items = new Dictionary<Guid, Item>();
    foreach (var kvp in currentSnapshot.Items)
        items.Add(kvp.Key, kvp.Value);
    items.Add(newWorkItemId, new Item(workItemName, workItemNotes));

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

    return newWorkItemId;
}

async Task CreateWorkItemFromJiraIssueAsync(JiraIssue jiraIssue, string connectionName)
{
    await CreateWorkItemFromJiraIssueAsyncInternal(jiraIssue, connectionName);
}

async Task ProcessWorklogsAsync(List<JiraWorklog> worklogs, string connectionName, Guid personId)
{
    if (jiraTransactor == null)
        throw new InvalidOperationException("Transactor not initialized. Extension may not be fully loaded.");

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

    var people = new Dictionary<Guid, Person>();
    foreach (var kvp in currentSnapshot.People)
        people.Add(kvp.Key, kvp.Value);

    var periods = new Dictionary<Guid, ItemPersonCorrelation<Period>>();
    foreach (var kvp in currentSnapshot.Periods)
        periods.Add(kvp.Key, kvp.Value);

    var periodsTiming = new List<Guid>();
    foreach (var item in currentSnapshot.PeriodsTiming)
        periodsTiming.Add(item);

    foreach (var worklog in worklogs)
    {
        var workItemId = await FindOrCreateWorkItemForJiraKeyAsync(worklog.IssueKey, worklog.IssueSummary, connectionName);

        if (!items.ContainsKey(workItemId))
        {
            var refreshedSnapshot = await jiraTransactor.SnapshotAsync();
            if (refreshedSnapshot.Items.ContainsKey(workItemId))
            {
                items.Add(workItemId, refreshedSnapshot.Items[workItemId]);
            }
        }

        var periodId = Guid.NewGuid();
        var endTime = worklog.Started.AddSeconds(worklog.TimeSpentSeconds);

        periods.Add(periodId, new ItemPersonCorrelation<Period>(
            workItemId,
            personId,
            new Period(endTime, worklog.Started, worklog.Comment)
        ));
    }

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
                ShowErrorDialog("Test Connection Failed", "Failed to connect to Jira server. Please check your connection details.", ex);
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

async void CreateWorkItemFromJiraClick(object sender, RoutedEventArgs e)
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

        var dialog = (Window)Extension.LoadUiElement("CreateWorkItemFromJiraDialog.xaml");
        var context = new CreateWorkItemFromJiraDialogContext();
        dialog.DataContext = context;

        var connectionComboBox = (ComboBox)dialog.FindName("connectionComboBox");
        var createButton = (Button)dialog.FindName("createButton");
        var cancelButton = (Button)dialog.FindName("cancelButton");
        var statusMessage = (TextBlock)dialog.FindName("statusMessage");

        connectionComboBox.ItemsSource = connections;
        if (connections.Count > 0)
            connectionComboBox.SelectedIndex = 0;

        createButton.Click += async (s, args) =>
        {
            try
            {
                createButton.IsEnabled = false;
                statusMessage.Text = "Fetching issue from Jira...";

                var selectedConnection = connectionComboBox.SelectedItem as JiraConnection;
                var issueKey = context.IssueKey?.Trim();

                if (selectedConnection == null)
                {
                    MessageDialog.Present(dialog, "Please select a Jira connection.", "Missing Connection", MessageBoxImage.Warning);
                    statusMessage.Text = "";
                    createButton.IsEnabled = true;
                    return;
                }

                if (string.IsNullOrEmpty(issueKey))
                {
                    MessageDialog.Present(dialog, "Please enter an issue key.", "Missing Issue Key", MessageBoxImage.Warning);
                    statusMessage.Text = "";
                    createButton.IsEnabled = true;
                    return;
                }

                var jiraIssue = await FetchJiraIssueAsync(
                    selectedConnection.ServerUrl,
                    selectedConnection.Email,
                    selectedConnection.ApiToken,
                    issueKey);

                statusMessage.Text = "Creating work item in Grindstone...";
                await CreateWorkItemFromJiraIssueAsync(jiraIssue, selectedConnection.Name);

                MessageDialog.Present(dialog,
                    $"Successfully created work item for {jiraIssue.Key}:\n{jiraIssue.Summary}",
                    "Success",
                    MessageBoxImage.Information);

                dialog.DialogResult = true;
            }
            catch (Exception ex)
            {
                statusMessage.Text = "";
                ShowErrorDialog("Error Creating Work Item", "Failed to create work item from Jira. Please verify the issue key and connection settings.", ex);
                createButton.IsEnabled = true;
            }
        };

        cancelButton.Click += (s, args) => dialog.DialogResult = false;

        dialog.ShowDialog();
    });
}

async void ImportWorklogsClick(object sender, RoutedEventArgs e)
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

        var dialog = (Window)Extension.LoadUiElement("ImportWorklogsDialog.xaml");

        var connectionComboBox = (ComboBox)dialog.FindName("connectionComboBox");
        var worklogDate = (DatePicker)dialog.FindName("worklogDate");
        var importButton = (Button)dialog.FindName("importButton");
        var cancelButton = (Button)dialog.FindName("cancelButton");
        var statusMessage = (TextBlock)dialog.FindName("statusMessage");

        connectionComboBox.ItemsSource = connections;
        if (connections.Count > 0)
            connectionComboBox.SelectedIndex = 0;

        worklogDate.SelectedDate = DateTime.Today;

        importButton.Click += async (s, args) =>
        {
            try
            {
                importButton.IsEnabled = false;
                statusMessage.Text = "Importing worklogs from Jira...";

                var selectedConnection = connectionComboBox.SelectedItem as JiraConnection;
                var selectedDate = worklogDate.SelectedDate;

                if (selectedConnection == null)
                {
                    MessageDialog.Present(dialog, "Please select a Jira connection.", "Missing Connection", MessageBoxImage.Warning);
                    statusMessage.Text = "";
                    importButton.IsEnabled = true;
                    return;
                }

                if (!selectedDate.HasValue)
                {
                    MessageDialog.Present(dialog, "Please select a date.", "Missing Date", MessageBoxImage.Warning);
                    statusMessage.Text = "";
                    importButton.IsEnabled = true;
                    return;
                }

                var worklogs = await ImportWorklogsFromJiraAsync(
                    selectedConnection.ServerUrl,
                    selectedConnection.Email,
                    selectedConnection.ApiToken,
                    selectedDate.Value);

                if (worklogs.Count == 0)
                {
                    MessageDialog.Present(dialog,
                        "No worklogs found for the selected date.",
                        "No Worklogs",
                        MessageBoxImage.Information);
                    statusMessage.Text = "";
                    importButton.IsEnabled = true;
                    return;
                }

                statusMessage.Text = $"Processing {worklogs.Count} worklog(s)...";

                var currentSnapshot = await jiraTransactor.SnapshotAsync();
                var currentPersonId = currentSnapshot.People.FirstOrDefault().Key;

                if (currentPersonId == Guid.Empty)
                {
                    MessageDialog.Present(dialog, "No person found in Grindstone. Please create a person first.", "Error", MessageBoxImage.Error);
                    statusMessage.Text = "";
                    importButton.IsEnabled = true;
                    return;
                }

                await ProcessWorklogsAsync(worklogs, selectedConnection.Name, currentPersonId);

                MessageDialog.Present(dialog,
                    $"Successfully imported {worklogs.Count} worklog(s) from Jira!",
                    "Success",
                    MessageBoxImage.Information);

                dialog.DialogResult = true;
            }
            catch (Exception ex)
            {
                statusMessage.Text = "";
                ShowErrorDialog("Error Importing Worklogs", "Failed to import worklogs from Jira. Please check the connection and date settings.", ex);
                importButton.IsEnabled = true;
            }
        };

        cancelButton.Click += (s, args) => dialog.DialogResult = false;

        dialog.ShowDialog();
    });
}

async void WorklogDashboardClick(object sender, RoutedEventArgs e)
{
    await Extension.OnUiThreadAsync(async () =>
    {
        if (jiraTransactor == null)
        {
            MessageDialog.Present(
                "The extension is not fully initialized. Please wait for Grindstone to finish loading.",
                "Not Ready",
                MessageBoxImage.Warning);
            return;
        }

        var dialog = (Window)Extension.LoadUiElement("WorklogDashboardDialog.xaml");

        var selectedDate = (DatePicker)dialog.FindName("selectedDate");
        var dateDisplayText = (TextBlock)dialog.FindName("dateDisplayText");
        var previousDayButton = (Button)dialog.FindName("previousDayButton");
        var nextDayButton = (Button)dialog.FindName("nextDayButton");
        var todayButton = (Button)dialog.FindName("todayButton");
        var syncToJiraButton = (Button)dialog.FindName("syncToJiraButton");
        var refreshButton = (Button)dialog.FindName("refreshButton");
        var worklogsGrid = (System.Windows.Controls.DataGrid)dialog.FindName("worklogsGrid");
        var totalTimeText = (TextBlock)dialog.FindName("totalTimeText");
        var entryCountText = (TextBlock)dialog.FindName("entryCountText");
        var closeButton = (Button)dialog.FindName("closeButton");
        var timeBarCanvas = (Canvas)dialog.FindName("timeBarCanvas");

        selectedDate.SelectedDate = DateTime.Today;

        Quantum.Entities.Frame currentSnapshot = null;

        Action renderTimeBar = () =>
        {
            if (currentSnapshot == null || !selectedDate.SelectedDate.HasValue || timeBarCanvas.ActualWidth == 0)
                return;

            var currentDate = selectedDate.SelectedDate.Value;
            var startOfDay = currentDate.Date;
            var endOfDay = startOfDay.AddDays(1);
            var currentPersonId = currentSnapshot.People.FirstOrDefault().Key;

            if (currentPersonId == Guid.Empty)
                return;

            timeBarCanvas.Children.Clear();

            // Calculate time range
            var minHour = 8;
            var maxHour = 18;
            var hasWorklogs = false;

            foreach (var periodEntry in currentSnapshot.Periods.Values)
            {
                if (periodEntry.PersonId != currentPersonId) continue;
                var p = periodEntry.CorrelatedEntity;
                if (p.Start.Date == currentDate.Date || p.End.Date == currentDate.Date)
                {
                    var effectiveStart = p.Start < startOfDay ? startOfDay : p.Start;
                    var effectiveEnd = p.End > endOfDay ? endOfDay : (p.End > DateTime.Now.AddYears(50) ? DateTime.Now : p.End);
                    
                    if (effectiveStart < effectiveEnd)
                    {
                        if (!hasWorklogs)
                        {
                            minHour = effectiveStart.Hour;
                            maxHour = effectiveEnd.Hour + (effectiveEnd.Minute > 0 ? 1 : 0);
                            hasWorklogs = true;
                        }
                        else
                        {
                            minHour = Math.Min(minHour, effectiveStart.Hour);
                            maxHour = Math.Max(maxHour, effectiveEnd.Hour + (effectiveEnd.Minute > 0 ? 1 : 0));
                        }
                    }
                }
            }

            if (maxHour <= minHour) maxHour = minHour + 1;
            
            // Add some padding if possible
            // minHour = Math.Max(0, minHour - 1);
            // maxHour = Math.Min(24, maxHour + 1);

            var totalHours = maxHour - minHour;

            // Draw hour markers
            for (int i = minHour; i <= maxHour; i++)
            {
                var x = ((i - minHour) / (double)totalHours) * timeBarCanvas.ActualWidth;
                var line = new System.Windows.Shapes.Line
                {
                    X1 = x,
                    Y1 = 0,
                    X2 = x,
                    Y2 = timeBarCanvas.ActualHeight,
                    Stroke = System.Windows.Media.Brushes.LightGray,
                    StrokeThickness = 1
                };
                timeBarCanvas.Children.Add(line);

                // Show label for every hour if space permits, or every 2 hours
                if (totalHours <= 12 || i % 2 == 0)
                {
                    var textBlock = new TextBlock
                    {
                        Text = $"{i:00}:00",
                        FontSize = 10,
                        Foreground = System.Windows.Media.Brushes.Gray
                    };
                    
                    // Center text
                    textBlock.Measure(new Size(double.PositiveInfinity, double.PositiveInfinity));
                    Canvas.SetLeft(textBlock, x - (textBlock.DesiredSize.Width / 2));
                    Canvas.SetTop(textBlock, timeBarCanvas.ActualHeight - 15);
                    timeBarCanvas.Children.Add(textBlock);
                }
            }

            // Draw worklog rectangles
            foreach (var periodEntry in currentSnapshot.Periods.Values)
            {
                if (periodEntry.PersonId != currentPersonId)
                    continue;

                var period = periodEntry.CorrelatedEntity;
                var periodStart = period.Start;
                var periodEnd = period.End;
                if (periodEnd > DateTime.Now.AddYears(50))
                    periodEnd = DateTime.Now;

                if (periodStart.Date == currentDate.Date ||
                    periodEnd.Date == currentDate.Date)
                {
                    var effectiveStart = periodStart < startOfDay ? startOfDay : periodStart;
                    var effectiveEnd = periodEnd > endOfDay ? endOfDay : periodEnd;

                    if (effectiveStart < endOfDay)
                    {
                        var totalMinutes = (int)(effectiveEnd - effectiveStart).TotalMinutes;
                        var startOffsetMinutes = (effectiveStart - startOfDay).TotalMinutes - (minHour * 60);
                        
                        if (startOffsetMinutes < 0) 
                        {
                            // Should not happen if logic is correct, but clamp just in case
                            totalMinutes += (int)startOffsetMinutes;
                            startOffsetMinutes = 0;
                        }

                        var left = (startOffsetMinutes / (totalHours * 60.0)) * timeBarCanvas.ActualWidth;
                        var width = (totalMinutes / (totalHours * 60.0)) * timeBarCanvas.ActualWidth;

                        if (width < 1) width = 1; // Minimum width visibility

                        var rect = new System.Windows.Shapes.Rectangle
                        {
                            Width = width,
                            Height = 20,
                            Fill = System.Windows.Media.Brushes.LightBlue,
                            Stroke = System.Windows.Media.Brushes.White,
                            StrokeThickness = 1,
                            ToolTip = $"{currentSnapshot.Items[periodEntry.ItemId].Name}\n{effectiveStart:HH:mm} - {effectiveEnd:HH:mm}"
                        };

                        Canvas.SetLeft(rect, left);
                        Canvas.SetTop(rect, 5);
                        timeBarCanvas.Children.Add(rect);
                    }
                }
            }
        };

        timeBarCanvas.SizeChanged += (s, args) => renderTimeBar();

        Action loadWorklogs = async () =>
        {
            try
            {
                if (!selectedDate.SelectedDate.HasValue)
                    return;

                var currentDate = selectedDate.SelectedDate.Value;
                dateDisplayText.Text = currentDate.ToString("dddd, MMMM d, yyyy");

                currentSnapshot = await jiraTransactor.SnapshotAsync();
                var currentPersonId = currentSnapshot.People.FirstOrDefault().Key;

                if (currentPersonId == Guid.Empty)
                    return;

                var startOfDay = currentDate.Date;
                var endOfDay = startOfDay.AddDays(1);

                var worklogs = new List<WorklogDisplayItem>();

                // Get Jira Key attribute ID
                var jiraKeyAttributeId = currentSnapshot.Attributes
                    .FirstOrDefault(kvp => kvp.Value.Name == "Jira Key")
                    .Key;

                // Get Jira Connection attribute ID
                var jiraConnectionAttributeId = currentSnapshot.Attributes
                    .FirstOrDefault(kvp => kvp.Value.Name == "Jira Connection")
                    .Key;

                foreach (var periodKvp in currentSnapshot.Periods)
                {
                    var periodId = periodKvp.Key;
                    var periodEntry = periodKvp.Value;

                    if (periodEntry.PersonId != currentPersonId)
                        continue;

                    var workItem = currentSnapshot.Items.ContainsKey(periodEntry.ItemId)
                        ? currentSnapshot.Items[periodEntry.ItemId]
                        : null;

                    if (workItem == null)
                        continue;

                    var period = periodEntry.CorrelatedEntity;
                    var periodStart = period.Start;
                    var periodEnd = period.End;
                    var isInProgress = periodEnd > DateTime.Now.AddYears(50); // Assuming MaxValue or far future means in progress
                    if (isInProgress)
                        periodEnd = DateTime.Now;

                    // Check if period overlaps with the selected day
                    if (periodStart.Date == currentDate.Date ||
                        periodEnd.Date == currentDate.Date)
                    {
                        var effectiveStart = periodStart < startOfDay ? startOfDay : periodStart;
                        var effectiveEnd = periodEnd > endOfDay ? endOfDay : periodEnd;

                        if (effectiveStart >= endOfDay)
                            continue;

                        var duration = effectiveEnd - effectiveStart;
                        var hours = (int)duration.TotalHours;
                        var minutes = duration.Minutes;
                        var durationText = hours > 0
                            ? $"{hours}h {minutes}m"
                            : $"{minutes}m";

                        worklogs.Add(new WorklogDisplayItem
                        {
                            StartTime = periodStart.ToString("yyyy-MM-dd HH:mm:ss"),
                            EndTime = isInProgress
                                ? "In Progress"
                                : periodEnd.ToString("yyyy-MM-dd HH:mm:ss"),
                            Duration = durationText,
                            WorkItemName = workItem.Name,
                            Notes = period.Notes ?? "",
                            PeriodId = periodId,
                            ItemId = periodEntry.ItemId,
                            JiraKey = jiraKeyAttributeId != Guid.Empty && currentSnapshot.AttributeValues.TryGetValue(new AttributeObjectCompositeKey(jiraKeyAttributeId, periodEntry.ItemId), out var jiraKeyVal)
                                ? jiraKeyVal as string
                                : "",
                            JiraConnection = jiraConnectionAttributeId != Guid.Empty && currentSnapshot.AttributeValues.TryGetValue(new AttributeObjectCompositeKey(jiraConnectionAttributeId, periodEntry.ItemId), out var jiraConnectionVal)
                                ? jiraConnectionVal as string
                                : ""
                        });
                    }
                }

                // Sort by start time
                worklogs = worklogs.OrderBy(w => w.StartTime).ToList();
                worklogsGrid.ItemsSource = worklogs;

                renderTimeBar();

                // Calculate totals
                var totalMinutesCalc = 0;
                foreach (var periodEntry in currentSnapshot.Periods.Values)
                {
                    if (periodEntry.PersonId != currentPersonId)
                        continue;

                    var period = periodEntry.CorrelatedEntity;
                    var periodStart = period.Start;
                    var periodEnd = period.End;
                    if (periodEnd > DateTime.Now.AddYears(50))
                        periodEnd = DateTime.Now;

                    if (periodStart.Date == currentDate.Date ||
                        periodEnd.Date == currentDate.Date)
                    {
                        var effectiveStart = periodStart < startOfDay ? startOfDay : periodStart;
                        var effectiveEnd = periodEnd > endOfDay ? endOfDay : periodEnd;

                        if (effectiveStart < endOfDay)
                        {
                            totalMinutesCalc += (int)(effectiveEnd - effectiveStart).TotalMinutes;
                        }
                    }
                }

                var totalHours = totalMinutesCalc / 60;
                var remainingMinutes = totalMinutesCalc % 60;
                totalTimeText.Text = $"Total Time: {totalHours}h {remainingMinutes}m";
                entryCountText.Text = $"Entries: {worklogs.Count}";
            }
            catch (Exception ex)
            {
                ShowErrorDialog("Error Loading Worklogs", "Failed to load worklogs from Grindstone.", ex);
            }
        };

        loadWorklogs();

        selectedDate.SelectedDateChanged += (s, args) => loadWorklogs();

        previousDayButton.Click += (s, args) =>
        {
            if (selectedDate.SelectedDate.HasValue)
            {
                selectedDate.SelectedDate = selectedDate.SelectedDate.Value.AddDays(-1);
            }
        };

        nextDayButton.Click += (s, args) =>
        {
            if (selectedDate.SelectedDate.HasValue)
            {
                selectedDate.SelectedDate = selectedDate.SelectedDate.Value.AddDays(1);
            }
        };

        todayButton.Click += (s, args) =>
        {
            selectedDate.SelectedDate = DateTime.Today;
        };

        refreshButton.Click += (s, args) => loadWorklogs();

        syncToJiraButton.Click += async (s, args) =>
        {
            try
            {
                syncToJiraButton.IsEnabled = false;

                var snapshot = await jiraTransactor.SnapshotAsync();
                var currentWorklogs = worklogsGrid.ItemsSource as List<WorklogDisplayItem>;

                if (currentWorklogs == null || currentWorklogs.Count == 0)
                {
                    MessageDialog.Present(dialog, "No worklogs to sync.", "No Worklogs", MessageBoxImage.Information);
                    syncToJiraButton.IsEnabled = true;
                    return;
                }

                // Get attribute IDs for Jira Key and Jira Connection
                var jiraKeyAttributeEntry = snapshot.Attributes.FirstOrDefault(kvp => kvp.Value.Name == "Jira Key");
                var jiraConnectionAttributeEntry = snapshot.Attributes.FirstOrDefault(kvp => kvp.Value.Name == "Jira Connection");

                if (jiraKeyAttributeEntry.Key == Guid.Empty || jiraConnectionAttributeEntry.Key == Guid.Empty)
                {
                    MessageDialog.Present(dialog,
                        "Required custom attributes 'Jira Key' and/or 'Jira Connection' not found. Please create work items from Jira first.",
                        "Missing Attributes",
                        MessageBoxImage.Warning);
                    syncToJiraButton.IsEnabled = true;
                    return;
                }

                // Check which worklogs have Jira Key and Jira Connection
                var worklogsToSync = new List<WorklogDisplayItem>();
                var worklogsSkipped = new List<WorklogDisplayItem>();

                foreach (var worklog in currentWorklogs)
                {
                    var jiraKeyValue = snapshot.AttributeValues
                        .Where(kvp => kvp.Key.AttributeId == jiraKeyAttributeEntry.Key && kvp.Key.ObjectId == worklog.ItemId)
                        .Select(kvp => kvp.Value as string)
                        .FirstOrDefault();

                    var jiraConnectionValue = snapshot.AttributeValues
                        .Where(kvp => kvp.Key.AttributeId == jiraConnectionAttributeEntry.Key && kvp.Key.ObjectId == worklog.ItemId)
                        .Select(kvp => kvp.Value as string)
                        .FirstOrDefault();

                    if (!string.IsNullOrEmpty(jiraKeyValue) && !string.IsNullOrEmpty(jiraConnectionValue))
                    {
                        worklogsToSync.Add(worklog);
                    }
                    else
                    {
                        worklogsSkipped.Add(worklog);
                    }
                }

                // If some worklogs don't have Jira Key or Jira Connection, show warning
                if (worklogsSkipped.Count > 0)
                {
                    var message = $"{worklogsSkipped.Count} out of {currentWorklogs.Count} worklog(s) do not have a Jira Key and/or Jira Connection assigned.\n\n" +
                                  "These worklogs will be skipped:\n";

                    var maxDisplay = Math.Min(5, worklogsSkipped.Count);
                    for (int i = 0; i < maxDisplay; i++)
                    {
                        message += $"- {worklogsSkipped[i].WorkItemName}\n";
                    }

                    if (worklogsSkipped.Count > maxDisplay)
                    {
                        message += $"... and {worklogsSkipped.Count - maxDisplay} more\n";
                    }

                    message += "\nDo you want to proceed with syncing the remaining worklogs?";

                    var result = MessageBox.Show(dialog, message, "Missing Jira Info", MessageBoxButton.YesNo, MessageBoxImage.Warning);

                    if (result != MessageBoxResult.Yes)
                    {
                        syncToJiraButton.IsEnabled = true;
                        return;
                    }
                }

                if (worklogsToSync.Count == 0)
                {
                    MessageDialog.Present(dialog, "No worklogs with Jira Key and Jira Connection found to sync.", "Nothing to Sync", MessageBoxImage.Information);
                    syncToJiraButton.IsEnabled = true;
                    return;
                }

                // Get Jira connections
                var connections = LoadJiraConnections();
                var connectionsByName = connections.ToDictionary(c => c.Name, c => c);

                // Sync each worklog
                var syncedCount = 0;
                var failedCount = 0;
                var errors = new List<string>();

                foreach (var worklog in worklogsToSync)
                {
                    try
                    {
                        var jiraKeyValue = snapshot.AttributeValues
                            .Where(kvp => kvp.Key.AttributeId == jiraKeyAttributeEntry.Key && kvp.Key.ObjectId == worklog.ItemId)
                            .Select(kvp => kvp.Value as string)
                            .FirstOrDefault();

                        var jiraConnectionValue = snapshot.AttributeValues
                            .Where(kvp => kvp.Key.AttributeId == jiraConnectionAttributeEntry.Key && kvp.Key.ObjectId == worklog.ItemId)
                            .Select(kvp => kvp.Value as string)
                            .FirstOrDefault();

                        if (string.IsNullOrEmpty(jiraConnectionValue) || !connectionsByName.ContainsKey(jiraConnectionValue))
                        {
                            errors.Add($"{worklog.WorkItemName}: No valid Jira connection found");
                            failedCount++;
                            continue;
                        }

                        var connection = connectionsByName[jiraConnectionValue];

                        // Get period details
                        var periodEntry = snapshot.Periods[worklog.PeriodId];
                        var period = periodEntry.CorrelatedEntity;
                        var timeSpentSeconds = (int)(period.End - period.Start).TotalSeconds;

                        if (timeSpentSeconds <= 0)
                        {
                            errors.Add($"{worklog.WorkItemName}: Invalid duration");
                            failedCount++;
                            continue;
                        }

                        await SyncWorklogToJiraAsync(
                            connection.ServerUrl,
                            connection.Email,
                            connection.ApiToken,
                            jiraKeyValue,
                            period.Start,
                            timeSpentSeconds,
                            period.Notes
                        );

                        syncedCount++;
                    }
                    catch (Exception ex)
                    {
                        errors.Add($"{worklog.WorkItemName}: {ex.Message}");
                        failedCount++;
                    }
                }

                // Show results
                var resultMessage = $"Sync completed!\n\nSuccessful: {syncedCount}\nFailed: {failedCount}";

                if (errors.Count > 0)
                {
                    resultMessage += "\n\nErrors:\n";
                    var maxErrors = Math.Min(5, errors.Count);
                    for (int i = 0; i < maxErrors; i++)
                    {
                        resultMessage += $"- {errors[i]}\n";
                    }
                    if (errors.Count > maxErrors)
                    {
                        resultMessage += $"... and {errors.Count - maxErrors} more errors";
                    }
                }

                MessageDialog.Present(dialog,
                    resultMessage,
                    "Sync Results",
                    failedCount > 0 ? MessageBoxImage.Warning : MessageBoxImage.Information);

                syncToJiraButton.IsEnabled = true;
            }
            catch (Exception ex)
            {
                ShowErrorDialog("Error Syncing to Jira", "Failed to sync worklogs to Jira.", ex);
                syncToJiraButton.IsEnabled = true;
            }
        };

        closeButton.Click += (s, args) => dialog.Close();

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

    var createWorkItemMenuItem = new RadMenuItem { Header = "Create Work Item from Jira" };
    createWorkItemMenuItem.Click += CreateWorkItemFromJiraClick;

    var importWorklogsMenuItem = new RadMenuItem { Header = "Import Worklogs" };
    importWorklogsMenuItem.Click += ImportWorklogsClick;

    var worklogDashboardMenuItem = new RadMenuItem { Header = "Worklog Dashboard" };
    worklogDashboardMenuItem.Click += WorklogDashboardClick;

    jiraMenuItem.Items.Add(manageConnectionsMenuItem);
    jiraMenuItem.Items.Add(createWorkItemMenuItem);
    jiraMenuItem.Items.Add(importWorklogsMenuItem);
    jiraMenuItem.Items.Add(worklogDashboardMenuItem);

    Extension.PostMessage(extensionsMenuExtensionId, jiraMenuItem);
});
