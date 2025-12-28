using Quantum;
using Quantum.Client.Windows;
using Quantum.Entities;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Text;
using System.IO;
using Microsoft.Win32;
using Telerik.Windows.Controls;

class CsvTransactor : Transactor
{
}

CsvTransactor csvTransactor;

void EngineStartingHandler(object sender, EngineStartingEventArgs e)
{
    csvTransactor = new CsvTransactor();
    e.AddTransactor(csvTransactor);
}

string EscapeCsv(string str)
{
    if (string.IsNullOrEmpty(str)) return "";
    // Replace actual line breaks with literal \r and \n sequences to ensure one line per record
    str = str.Replace("\r\n", "\\r\\n").Replace("\n", "\\n").Replace("\r", "\\r").Trim();
    
    str = str.Replace("\"", "\"\"");
    if (str.Contains(",") || str.Contains("\""))
    {
        return $"\"{str}\"";
    }
    return str;
}

async void ExportToCsvClick(object sender, RoutedEventArgs e)
{
    await Extension.OnUiThreadAsync(async () =>
    {
        if (csvTransactor == null)
        {
             MessageDialog.Present("Extension not ready.", "Not Ready", MessageBoxImage.Warning);
             return;
        }

        var dialog = (Window)Extension.LoadUiElement("CSVExportDialog.xaml");
        var startDatePicker = (DatePicker)dialog.FindName("startDatePicker");
        var endDatePicker = (DatePicker)dialog.FindName("endDatePicker");
        var exportButton = (Button)dialog.FindName("exportButton");
        var cancelButton = (Button)dialog.FindName("cancelButton");

        var includeOriginalNameCheckbox = (CheckBox)dialog.FindName("includeOriginalNameCheckbox");

        startDatePicker.SelectedDate = DateTime.Today.AddDays(-7);
        endDatePicker.SelectedDate = DateTime.Today;

        exportButton.Click += async (s, args) =>
        {
            var start = startDatePicker.SelectedDate ?? DateTime.MinValue;
            var end = (endDatePicker.SelectedDate ?? DateTime.MaxValue).Date.AddDays(1).AddSeconds(-1);
            var includeOriginalName = includeOriginalNameCheckbox.IsChecked ?? true;

            if (start > end)
            {
                MessageDialog.Present("Start date must be before end date.", "Invalid Range", MessageBoxImage.Error);
                return;
            }

            var saveDialog = new SaveFileDialog
            {
                Filter = "CSV files (*.csv)|*.csv|All files (*.*)|*.*",
                FileName = $"Grindstone_Export_{start:yyyyMMdd}_{endDatePicker.SelectedDate:yyyyMMdd}.csv"
            };

            if (saveDialog.ShowDialog() == true)
            {
                try
                {
                    var snapshot = await csvTransactor.SnapshotAsync();
                    
                    // Get Jira Key attribute ID
                    var jiraKeyAttributeId = snapshot.Attributes
                        .FirstOrDefault(kvp => kvp.Value.Name == "Jira Key")
                        .Key;

                    var csv = new StringBuilder();
                    var headers = new List<string> { "Start of timeslice", "End of timeslice", "timeslice notes", "WorkItem", "Jira key" };
                    if (includeOriginalName) headers.Add("Original WorkItem");
                    csv.AppendLine(string.Join(",", headers));

                    foreach (var periodEntry in snapshot.Periods.Values)
                    {
                        var period = periodEntry.CorrelatedEntity;
                        var pStart = period.Start;
                        var pEnd = (period.End > DateTime.Now.AddYears(50) ? DateTime.Now : period.End);

                        // Check if period overlaps with selected range
                        if (pStart >= end || pEnd <= start) continue;

                        var originalItemName = snapshot.Items.ContainsKey(periodEntry.ItemId) ? snapshot.Items[periodEntry.ItemId].Name : "Unknown";
                        var itemName = originalItemName;
                        
                        string jiraKey = "";
                        if (jiraKeyAttributeId != Guid.Empty && snapshot.AttributeValues.TryGetValue(new AttributeObjectCompositeKey(jiraKeyAttributeId, periodEntry.ItemId), out var jiraKeyVal))
                        {
                            jiraKey = jiraKeyVal as string;
                        }

                        // Extraction and Cleaning logic
                        // Try to parse name: 'PFTI-092 - Create Payment' or '[PFTI-092] Verlof'
                        // Regular expression to find a Jira key at the start, with optional brackets, 
                        // followed by a separator (space-dash-space OR just space) and the rest of the text.
                        var match = System.Text.RegularExpressions.Regex.Match(originalItemName, @"^\[?([A-Z0-9]+-\d+)\]?(?:\s*-\s*|\s+)(.*)$");
                        
                        if (match.Success)
                        {
                            var extractedKey = match.Groups[1].Value;
                            if (string.IsNullOrWhiteSpace(jiraKey))
                            {
                                jiraKey = extractedKey;
                            }
                            
                            // The description should be the remainder regardless of whether the key matches the attribute
                            itemName = match.Groups[2].Value.Trim();
                        }

                        var row = new List<string>
                        {
                            EscapeCsv(pStart.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")),
                            EscapeCsv(pEnd.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")),
                            EscapeCsv(period.Notes),
                            EscapeCsv(itemName),
                            EscapeCsv(jiraKey)
                        };
                        if (includeOriginalName) row.Add(EscapeCsv(originalItemName));

                        csv.AppendLine(string.Join(",", row));
                    }

                    File.WriteAllText(saveDialog.FileName, csv.ToString(), Encoding.UTF8);
                    MessageDialog.Present("Export completed successfully.", "Success", MessageBoxImage.Information);
                    dialog.Close();
                }
                catch (Exception ex)
                {
                    MessageDialog.Present($"Export failed: {ex.Message}", "Error", MessageBoxImage.Error);
                }
            }
        };

        cancelButton.Click += (s, args) => dialog.Close();

        dialog.ShowDialog();
    });
}

var extensionsMenuExtensionId = Guid.Parse("{27F65593-7235-4108-B5D9-F0DE417D8536}");

Extension.EngineStarting += EngineStartingHandler;

await Extension.OnUiThreadAsync(() =>
{
    var exportMenuItem = new RadMenuItem { Header = "Export to CSV..." };
    exportMenuItem.Click += ExportToCsvClick;

    Extension.PostMessage(extensionsMenuExtensionId, exportMenuItem);
});
