using System;
using System.Collections.ObjectModel;
using System.IO;
using System.Linq;
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
                    }
                    else
                    {
                        Events.Add(new Event { EventFile = openFileDialog.FileName });
                    }
                }
            }
        }

        private void RemoveEventFileButton_Click(object sender, RoutedEventArgs e)
        {
            var button = sender as Button;
            if (button != null)
            {
                var eventItem = button.DataContext as Event;
                if (eventItem != null)
                {
                    Events.Remove(eventItem);
                }
            }
        }

        private void SynchronizeButton_Click(object sender, RoutedEventArgs e)
        {
            string filePath = FilePathTextBox.Text;

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

            try
            {
                using (var fileStream = new FileStream(filePath, FileMode.Open, FileAccess.Read))
                {
                    IWorkbook workbook = new XSSFWorkbook(fileStream);
                    ISheet sheet = workbook.GetSheet("節目贊助");
                    if (sheet == null)
                    {
                        MessageBox.Show("Sheet '節目贊助' not found.");
                        return;
                    }

                    IRow row = sheet.GetRow(2); // F3 is the third row (index 2)
                    if (row == null)
                    {
                        MessageBox.Show("Row 3 not found.");
                        return;
                    }

                    var existingEventNames = new Dictionary<string, int>();
                    for (int col = 5; col < row.LastCellNum; col++) // F is the sixth column (index 5)
                    {
                        ICell cell = row.GetCell(col);
                        if (cell != null)
                        {
                            string eventName = cell.ToString().Replace("\r\n", string.Empty).Replace("\n", string.Empty);
                            if (!string.IsNullOrEmpty(eventName))
                            {
                                existingEventNames[eventName] = col;
                            }
                        }
                    }

                    foreach (var eventItem in Events)
                    {
                        string eventFileName = Path.GetFileNameWithoutExtension(eventItem.EventFile);
                        string eventFileNamePrefix = eventFileName.Split('_')[0]; // Get substring before underscore
                        if (existingEventNames.TryGetValue(eventFileNamePrefix, out int colIndex))
                        {
                            // Process existing event donation amount cells
                            ProcessDonationAmountCells(sheet, colIndex);
                        }
                        else
                        {
                            // Add new event name to the overview file
                            int newColIndex = row.LastCellNum;
                            ICell newCell = row.CreateCell(newColIndex);
                            newCell.SetCellValue(eventFileNamePrefix);
                            existingEventNames[eventFileNamePrefix] = newColIndex;

                            // Process new event donation amount cells
                            ProcessDonationAmountCells(sheet, newColIndex);
                        }
                    }

                    // Save the updated workbook
                    using (var outputStream = new FileStream(filePath, FileMode.Create, FileAccess.Write))
                    {
                        workbook.Write(outputStream);
                    }

                    MessageBox.Show("Overview file processed and updated.");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"An error occurred while processing the overview file: {ex.Message}");
            }
        }

        private void ProcessDonationAmountCells(ISheet sheet, int colIndex)
        {
            // Locate event donation amount cells (20 cells below event name)
            for (int rowIndex = 3; rowIndex <= 23; rowIndex++) // 20 cells below row 3 is row 4 to row 24 (index 3 to 23)
            {
                IRow donationRow = sheet.GetRow(rowIndex);
                if (donationRow != null)
                {
                    ICell donationCell = donationRow.GetCell(colIndex);
                    if (donationCell != null)
                    {
                        // Process the donation amount cell as needed
                        // For example, you can read the value or update it
                        string donationAmount = donationCell.ToString();

                        // Look for board member name under cell '董事會成員' in column A
                        ICell boardMemberCell = donationRow.GetCell(0); // Column A is index 0
                        if (boardMemberCell != null)
                        {
                            string boardMemberName = boardMemberCell.ToString();
                            // Do something with the boardMemberName and donationAmount
                        }
                    }
                }
            }
        }
    }

    public class Event
    {
        public string EventFile { get; set; }
    }
}
