using Quantum;
using Quantum.Client.Windows;
using Quantum.Entities;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using Telerik.Windows.Controls;
using System.Data;
using System.Windows.Media;

class JiraTransactor : Transactor
{
}

JiraTransactor jiraTransactor;

void EngineStartingHandler(object sender, EngineStartingEventArgs e)
{
    jiraTransactor = new JiraTransactor();
    e.AddTransactor(jiraTransactor);
}

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

async void MonthlyOverviewClick(object sender, RoutedEventArgs e)
{
    await Extension.OnUiThreadAsync(async () =>
    {
        if (jiraTransactor == null)
        {
             MessageDialog.Present("Extension not ready.", "Not Ready", MessageBoxImage.Warning);
             return;
        }

        var dialog = (Window)Extension.LoadUiElement("MonthlyOverviewDialog.xaml");
        var monthComboBox = (ComboBox)dialog.FindName("monthComboBox");
        var yearComboBox = (ComboBox)dialog.FindName("yearComboBox");
        var refreshButton = (Button)dialog.FindName("refreshButton");
        var overviewGrid = (DataGrid)dialog.FindName("overviewGrid");
        var closeButton = (Button)dialog.FindName("closeButton");
        var statusText = (TextBlock)dialog.FindName("statusText");
        var totalWorkdaysText = (TextBlock)dialog.FindName("totalWorkdaysText");
        var averageHoursText = (TextBlock)dialog.FindName("averageHoursText");
        var overtimeText = (TextBlock)dialog.FindName("overtimeText");

        // Populate Month/Year
        var months = System.Globalization.CultureInfo.CurrentCulture.DateTimeFormat.MonthNames.Where(m => !string.IsNullOrEmpty(m)).ToList();
        monthComboBox.ItemsSource = months;
        monthComboBox.SelectedIndex = DateTime.Today.Month - 1;

        var currentYear = DateTime.Today.Year;
        var years = Enumerable.Range(currentYear - 5, 6).ToList();
        yearComboBox.ItemsSource = years;
        yearComboBox.SelectedItem = currentYear;

        // Handle Column Generation
        overviewGrid.AutoGeneratingColumn += (s, args) =>
        {
            // Common Header Style
            var headerStyle = new Style(typeof(DataGridColumnHeader));
            headerStyle.Setters.Add(new Setter(Control.BackgroundProperty, new SolidColorBrush((Color)ColorConverter.ConvertFromString("#61a9e0"))));
            headerStyle.Setters.Add(new Setter(Control.ForegroundProperty, Brushes.White));
            headerStyle.Setters.Add(new Setter(Control.FontWeightProperty, FontWeights.Bold));
            headerStyle.Setters.Add(new Setter(Control.HorizontalContentAlignmentProperty, HorizontalAlignment.Center));
            args.Column.HeaderStyle = headerStyle;

            if (args.PropertyName == "Jira Key")
            {
                args.Column.Width = DataGridLength.Auto;
            }
            else if (args.PropertyName == "Work Item")
            {
                args.Column.Width = new DataGridLength(1, DataGridLengthUnitType.Star);
                args.Column.MinWidth = 200;
            }
            else if (args.PropertyName == "Total")
            {
                args.Column.Width = DataGridLength.Auto;
                
                // Cell Style
                var style = new Style(typeof(TextBlock));
                style.Setters.Add(new Setter(TextBlock.FontWeightProperty, FontWeights.Bold));
                style.Setters.Add(new Setter(TextBlock.TextAlignmentProperty, TextAlignment.Center));
                ((DataGridTextColumn)args.Column).ElementStyle = style;
            }
            else
            {
                args.Column.Width = new DataGridLength(40);
                
                // Center Cell Content
                var cellStyle = new Style(typeof(TextBlock));
                cellStyle.Setters.Add(new Setter(TextBlock.TextAlignmentProperty, TextAlignment.Center));
                
                // Check for Weekend
                if (int.TryParse(args.PropertyName, out int day))
                {
                    var selectedMonthIndex = monthComboBox.SelectedIndex;
                    var selectedYear = (int)yearComboBox.SelectedItem;
                    if (selectedMonthIndex >= 0)
                    {
                        try
                        {
                            var date = new DateTime(selectedYear, selectedMonthIndex + 1, day);
                            if (date.DayOfWeek == DayOfWeek.Saturday || date.DayOfWeek == DayOfWeek.Sunday)
                            {
                                cellStyle.Setters.Add(new Setter(TextBlock.BackgroundProperty, Brushes.LightGray));
                            }
                        }
                        catch {}
                    }
                }

                ((DataGridTextColumn)args.Column).ElementStyle = cellStyle;
            }
        };

        // Loading Row Event for Total Row Color
        overviewGrid.LoadingRow += (s, args) =>
        {
            var row = args.Row;
            var item = row.Item as DataRowView;
            if (item != null && item["Work Item"].ToString() == "TOTAL")
            {
                row.Background = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#61a9e0"));
                row.Foreground = Brushes.White;
                row.FontWeight = FontWeights.Bold;
            }
            else
            {
                row.Background = Brushes.White;
                row.Foreground = Brushes.Black;
                row.FontWeight = FontWeights.Normal;
            }
        };

        Action loadData = async () =>
        {
            try 
            {
                statusText.Text = "Loading...";
                refreshButton.IsEnabled = false;
                overviewGrid.ItemsSource = null;
                
                var selectedMonthIndex = monthComboBox.SelectedIndex;
                var selectedYear = (int)yearComboBox.SelectedItem;
                
                if (selectedMonthIndex < 0) return;

                var daysInMonth = DateTime.DaysInMonth(selectedYear, selectedMonthIndex + 1);
                var startDate = new DateTime(selectedYear, selectedMonthIndex + 1, 1);
                var endDate = startDate.AddMonths(1);

                var snapshot = await jiraTransactor.SnapshotAsync();
                var personId = snapshot.People.FirstOrDefault().Key;

                if (personId == Guid.Empty) return;

                // Get Jira Key attribute ID
                var jiraKeyAttributeId = snapshot.Attributes
                    .FirstOrDefault(kvp => kvp.Value.Name == "Jira Key")
                    .Key;

                // Calculate Workdays
                int workdays = 0;
                for (int i = 1; i <= daysInMonth; i++)
                {
                    var date = new DateTime(selectedYear, selectedMonthIndex + 1, i);
                    if (date.DayOfWeek != DayOfWeek.Saturday && date.DayOfWeek != DayOfWeek.Sunday)
                    {
                        workdays++;
                    }
                }
                totalWorkdaysText.Text = workdays.ToString();

                // Prepare DataTable
                var dt = new DataTable();
                dt.Columns.Add("Jira Key", typeof(string));
                dt.Columns.Add("Work Item", typeof(string));
                for (int i = 1; i <= daysInMonth; i++)
                {
                    dt.Columns.Add(i.ToString(), typeof(double));
                }
                dt.Columns.Add("Total", typeof(double));

                // Aggregate Data
                var workItemHours = new Dictionary<Guid, Dictionary<int, double>>();
                var workItemNames = new Dictionary<Guid, string>();
                var workItemJiraKeys = new Dictionary<Guid, string>();
                var dailyTotals = new Dictionary<int, double>();
                double grandTotal = 0;

                foreach (var periodEntry in snapshot.Periods.Values)
                {
                    if (periodEntry.PersonId != personId) continue;
                    
                    var period = periodEntry.CorrelatedEntity;
                    // Intersect period with month
                    var pStart = period.Start < startDate ? startDate : period.Start;
                    var pEnd = period.End > endDate ? endDate : (period.End > DateTime.Now.AddYears(50) ? DateTime.Now : period.End);

                    if (pStart >= pEnd) continue;
                    if (pStart >= endDate || pEnd <= startDate) continue;

                    // Split by day
                    var current = pStart;
                    while (current < pEnd)
                    {
                        var nextDay = current.Date.AddDays(1);
                        var segmentEnd = nextDay < pEnd ? nextDay : pEnd;
                        
                        var duration = (segmentEnd - current).TotalHours;
                        var day = current.Day;

                        if (!workItemHours.ContainsKey(periodEntry.ItemId))
                        {
                            workItemHours[periodEntry.ItemId] = new Dictionary<int, double>();
                            var item = snapshot.Items.ContainsKey(periodEntry.ItemId) ? snapshot.Items[periodEntry.ItemId] : null;
                            workItemNames[periodEntry.ItemId] = item?.Name ?? "Unknown";
                            
                            // Get Jira Key
                            string jiraKey = "";
                            if (jiraKeyAttributeId != Guid.Empty && snapshot.AttributeValues.TryGetValue(new AttributeObjectCompositeKey(jiraKeyAttributeId, periodEntry.ItemId), out var jiraKeyVal))
                            {
                                jiraKey = jiraKeyVal as string;
                            }
                            workItemJiraKeys[periodEntry.ItemId] = jiraKey;
                        }

                        if (!workItemHours[periodEntry.ItemId].ContainsKey(day))
                            workItemHours[periodEntry.ItemId][day] = 0;
                        
                        workItemHours[periodEntry.ItemId][day] += duration;

                        if (!dailyTotals.ContainsKey(day)) dailyTotals[day] = 0;
                        dailyTotals[day] += duration;
                        grandTotal += duration;

                        current = nextDay;
                    }
                }

                // Populate Rows
                foreach (var itemId in workItemHours.Keys)
                {
                    var row = dt.NewRow();
                    row["Jira Key"] = workItemJiraKeys[itemId];
                    row["Work Item"] = workItemNames[itemId];
                    double rowTotal = 0;
                    foreach (var day in workItemHours[itemId].Keys)
                    {
                        var hours = workItemHours[itemId][day];
                        row[day.ToString()] = Math.Round(hours, 2);
                        rowTotal += hours;
                    }
                    row["Total"] = Math.Round(rowTotal, 2);
                    dt.Rows.Add(row);
                }

                // Add Summary Row
                var summaryRow = dt.NewRow();
                summaryRow["Jira Key"] = "";
                summaryRow["Work Item"] = "TOTAL";
                foreach (var day in dailyTotals.Keys)
                {
                    summaryRow[day.ToString()] = Math.Round(dailyTotals[day], 2);
                }
                summaryRow["Total"] = Math.Round(grandTotal, 2);
                dt.Rows.Add(summaryRow);

                // Update Stats
                int daysWithWorklogsOnWeekdays = 0;
                foreach (var day in dailyTotals.Keys)
                {
                    if (dailyTotals[day] > 0)
                    {
                        var date = new DateTime(selectedYear, selectedMonthIndex + 1, day);
                        if (date.DayOfWeek != DayOfWeek.Saturday && date.DayOfWeek != DayOfWeek.Sunday)
                        {
                            daysWithWorklogsOnWeekdays++;
                        }
                    }
                }

                double average = daysWithWorklogsOnWeekdays > 0 ? grandTotal / daysWithWorklogsOnWeekdays : 0;
                averageHoursText.Text = Math.Round(average, 2).ToString("F2");

                double overtime = grandTotal - (workdays * 8);
                overtimeText.Text = Math.Round(overtime, 2).ToString("F2");
                overtimeText.Foreground = overtime < 0 ? Brushes.Red : Brushes.Green;

                overviewGrid.ItemsSource = dt.DefaultView;
                statusText.Text = "";
            }
            catch (Exception ex)
            {
                statusText.Text = "Error";
                ShowErrorDialog("Error", "Failed to load data", ex);
            }
            finally
            {
                refreshButton.IsEnabled = true;
            }
        };

        refreshButton.Click += (s, args) => loadData();
        closeButton.Click += (s, args) => dialog.Close();

        loadData();
        dialog.ShowDialog();
    });
}

var extensionsMenuExtensionId = Guid.Parse("{27F65593-7235-4108-B5D9-F0DE417D8536}");

Extension.EngineStarting += EngineStartingHandler;

await Extension.OnUiThreadAsync(() =>
{
    var monthlyOverviewMenuItem = new RadMenuItem { Header = "Monthly Overview" };
    monthlyOverviewMenuItem.Click += MonthlyOverviewClick;

    Extension.PostMessage(extensionsMenuExtensionId, monthlyOverviewMenuItem);
});
