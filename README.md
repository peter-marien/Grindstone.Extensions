# Grindstone Jira Extension

This extension integrates [Grindstone 4](https://www.epiforge.com/grindstone/) with [Jira](https://www.atlassian.com/software/jira), allowing you to seamlessly manage your time tracking and synchronization with Jira issues.

## Features

### 1. Manage Jira Connections
-   Configure multiple Jira server connections.
-   Securely store API tokens and credentials.
-   Test connections to ensure valid configuration.

### 2. Create Work Item from Jira
-   Quickly create a new Grindstone work item based on a Jira issue key.
-   Automatically fetches the issue summary and populates the work item name.
-   Links the work item to the Jira issue using custom attributes ('Jira Key' and 'Jira Connection').

### 3. Import Worklogs
-   Import existing worklogs from Jira for a specific date.
-   Useful for keeping your local Grindstone data in sync with work logged directly in Jira.

### 4. Worklog Dashboard
-   View a daily summary of your worklogs.
-   **Dynamic Time-Bar**: Visualizes your work distribution throughout the day, automatically adjusting the range based on your earliest and latest logs.
-   **Enhanced DataGrid**: Includes columns for 'Jira Key' and 'Jira Connection' for better visibility.
-   **Sync to Jira**: Upload your local Grindstone time tracking data to Jira.
    -   Ensures work items are properly linked (requires both 'Jira Key' and 'Jira Connection').
    -   Validates data before syncing.

### 5. Monthly Overview
-   **Matrix View**: Displays a comprehensive overview of hours worked per work item for each day of the selected month.
-   **Statistics**:
    -   **Total Workdays**: Counts the number of weekdays (Mon-Fri) in the month.
    -   **Average Hours/WorkDay**: Calculates the average hours worked on weekdays where work was logged.
    -   **Overtime**: Tracks your overtime balance based on a standard 8-hour workday (Total Hours - (Total Weekdays * 8)).
-   **Summary**: Includes a "Total" row summing up hours for each day and a grand total for the month.

## Installation & Deployment

1.  Ensure Grindstone 4 is installed.
2.  Run the included PowerShell script to deploy the extension:
    ```powershell
    .\DeployExtensions.ps1
    ```
    This script copies the extension files to your `%APPDATA%\Grindstone 4\Extensions` directory.
3.  Restart Grindstone 4 to load the extension.

## Requirements

-   Grindstone 4
-   Jira Cloud or Server account
-   Jira API Token (for Cloud) or Password (for Server)
