using Quantum.Client.Windows;
using System;
using System.Windows;
using Telerik.Windows.Controls;

var extensionsMenuExtensionId = Guid.Parse("{27F65593-7235-4108-B5D9-F0DE417D8536}");

void FetchFromJiraClick(object sender, RoutedEventArgs e)
{
    MessageDialog.Present(
        "Jira fetch functionality will be implemented here.",
        "Fetch from Jira",
        MessageBoxImage.Information
    );
}

await Extension.OnUiThreadAsync(() =>
{
    var jiraMenuItem = new RadMenuItem { Header = "Jira Integration" };

    var fetchMenuItem = new RadMenuItem { Header = "Fetch from Jira" };
    fetchMenuItem.Click += FetchFromJiraClick;

    jiraMenuItem.Items.Add(fetchMenuItem);

    Extension.PostMessage(extensionsMenuExtensionId, jiraMenuItem);
});
