using System;
using System.Collections.ObjectModel;
using System.IO;
using System.Windows;
using System.Windows.Controls;
using Microsoft.Win32;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

namespace sync_donations_bm
{
    public partial class MainWindow : Window
    {
        public ObservableCollection<Event> Events { get; set; }

        public MainWindow()
        {
            InitializeComponent();
            Events = new ObservableCollection<Event>();
            EventsDataGrid.ItemsSource = Events;
        }

        private void BrowseButton_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog
            {
                Filter = "Excel Files|*.xlsx;*.xls",
                Title = "Select an Excel File"
            };

            if (openFileDialog.ShowDialog() == true)
            {
                FilePathTextBox.Text = openFileDialog.FileName;
                string filePath = openFileDialog.FileName;

                if (string.IsNullOrWhiteSpace(filePath))
                {
                    MessageBox.Show("Please enter a file path.");
                    return;
                }

                if (!File.Exists(filePath))
                {
                    MessageBox.Show("File not found.");
                    return;
                }

                using (var fileStream = new FileStream(filePath, FileMode.Open, FileAccess.Read))
                {
                    IWorkbook workbook = new XSSFWorkbook(fileStream);
                    ISheet sheet = workbook.GetSheet("節目贊助");
                    if (sheet == null)
                    {
                        MessageBox.Show("Sheet '節目贊助' not found.");
                        return;
                    }

                    Events.Clear();
                    IRow row = sheet.GetRow(2); // F3 is the third row (index 2)
                    if (row != null)
                    {
                        for (int col = 5; col < row.LastCellNum; col++) // F is the sixth column (index 5)
                        {
                            ICell cell = row.GetCell(col);
                            if (cell != null)
                            {
                                string eventName = cell.ToString().Replace("\r\n", string.Empty).Replace("\n", string.Empty);
                                if (!string.IsNullOrEmpty(eventName))
                                {
                                    Events.Add(new Event { Name = eventName });
                                }
                            }
                        }
                    }

                    // Process the events as needed
                    MessageBox.Show($"Found {Events.Count} events.");
                }
            }
        }

        private void BrowseEventFileButton_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog
            {
                Filter = "Excel Files|*.xlsx;*.xls",
                Title = "Select an Event Excel File"
            };

            if (openFileDialog.ShowDialog() == true)
            {
                var button = sender as Button;
                if (button != null)
                {
                    var eventItem = button.DataContext as Event;
                    if (eventItem != null)
                    {
                        eventItem.EventFile = openFileDialog.FileName;
                        // Update the TextBox in the DataGrid
                        var parent = button.Parent as StackPanel;
                        if (parent != null)
                        {
                            var textBox = parent.Children[0] as TextBox;
                            if (textBox != null)
                            {
                                textBox.Text = openFileDialog.FileName;
                            }
                        }
                    }
                }
            }
        }

        private void SynchronizeButton_Click(object sender, RoutedEventArgs e)
        {
            foreach (var eventItem in Events)
            {
                if (string.IsNullOrWhiteSpace(eventItem.EventFile) || !File.Exists(eventItem.EventFile))
                {
                    MessageBox.Show($"Event file for '{eventItem.Name}' is missing or does not exist.");
                    return;
                }

                string eventFileName = Path.GetFileNameWithoutExtension(eventItem.EventFile);
                if (!eventFileName.Contains(eventItem.Name))
                {
                    MessageBox.Show($"Event file name '{eventFileName}' does not contain the event name '{eventItem.Name}'.");
                    return;
                }
            }

            // Further processing if needed
            MessageBox.Show("All event files are valid.");
        }
    }

    public class Event
    {
        public string Name { get; set; }
        public string EventFile { get; set; }
    }
}
