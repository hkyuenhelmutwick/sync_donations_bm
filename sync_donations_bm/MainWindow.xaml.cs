﻿using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.IO;
using System.Linq;
using System.Text.Json;
using System.Windows;
using System.Windows.Controls;
using Microsoft.Win32;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

namespace sync_donations_bm
{
    public partial class MainWindow : Window
    {
        private static readonly string JsonFilePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "events.json");
        private static readonly string ProjectJsonFilePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "..", "..", "..", "events.json");
        public ObservableCollection<Event> Events { get; set; }
        public Overview Overview { get; set; }

        public MainWindow()
        {
            InitializeComponent();
            Events = new ObservableCollection<Event>();
            EventsDataGrid.ItemsSource = Events;
            LoadEventsFromJson();
            if (Overview != null && !string.IsNullOrWhiteSpace(Overview.OverviewFilePath))
            {
                FilePathTextBox.Text = Overview.OverviewFilePath;
            }
        }

        private void LoadEventsFromJson()
        {
            if (File.Exists(JsonFilePath))
            {
                var json = File.ReadAllText(JsonFilePath);
                var data = JsonSerializer.Deserialize<OverviewData>(json);
                if (data != null)
                {
                    Overview = data.Overview;
                    foreach (var eventItem in data.Events)
                    {
                        Events.Add(eventItem);
                    }
                }
            }
        }

        private void SaveEventsToJson()
        {
            var data = new OverviewData
            {
                Overview = Overview,
                Events = new List<Event>(Events)
            };
            var json = JsonSerializer.Serialize(data);
            File.WriteAllText(JsonFilePath, json);
            File.WriteAllText(ProjectJsonFilePath, json); // Update the project directory as well
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

                // Save the overview file path to events.json
                Overview = new Overview { OverviewFilePath = openFileDialog.FileName };
                SaveEventsToJson();
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
                    SaveEventsToJson();
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
                    SaveEventsToJson();
                }
            }
        }

        private void SynchronizeButton_Click(object sender, RoutedEventArgs e)
        {
            if (Overview == null || string.IsNullOrWhiteSpace(Overview.OverviewFilePath))
            {
                MessageBox.Show("Please select an overview file.");
                return;
            }

            string filePath = Overview.OverviewFilePath;

            if (!File.Exists(filePath))
            {
                if (IsProductionEnvironment())
                {
                    // In production, continue without prompting an error
                    return;
                }
                else
                {
                    MessageBox.Show("File not found.");
                    return;
                }
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

                    // Collect existing event names and their column indices
                    var existingEventNames = CollectExistingEventNames(row);

                    // Process each event file
                    foreach (var eventItem in Events)
                    {
                        ProcessEventFile(sheet, row, existingEventNames, eventItem);
                    }

                    // Save the updated workbook
                    using (var outputStream = new FileStream(filePath, FileMode.Create, FileAccess.Write))
                    {
                        workbook.Write(outputStream);
                    }

                    MessageBox.Show("Overview file processed and updated.");
                    SaveEventsToJson();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"An error occurred while processing the overview file: {ex.Message}");
            }
        }

        private bool IsProductionEnvironment()
        {
            // Check if only the events.json file exists in the bin directory
            return File.Exists(JsonFilePath) && !File.Exists(ProjectJsonFilePath);
        }

        private Dictionary<string, int> CollectExistingEventNames(IRow row)
        {
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
            return existingEventNames;
        }

        private void ProcessEventFile(ISheet sheet, IRow row, Dictionary<string, int> existingEventNames, Event eventItem)
        {
            string eventFileName = Path.GetFileNameWithoutExtension(eventItem.EventFile);
            string eventFileNamePrefix = eventFileName.Split('_')[0]; // Get substring before underscore
            int colIndex;
            if (!existingEventNames.TryGetValue(eventFileNamePrefix, out colIndex))
            {
                // Add new event name to the overview file
                colIndex = row.LastCellNum;
                ICell newCell = row.CreateCell(colIndex);
                newCell.SetCellValue(eventFileNamePrefix);
                existingEventNames[eventFileNamePrefix] = colIndex;

                // Create 21 cells under the new event name cell
                for (int rowIndex = 3; rowIndex <= 23; rowIndex++) // 21 cells below row 3 is row 4 to row 24 (index 3 to 23)
                {
                    IRow donationRow = sheet.GetRow(rowIndex) ?? sheet.CreateRow(rowIndex);
                    donationRow.CreateCell(colIndex);
                }
            }

            // Process event donation amount cells
            ProcessDonationAmountCells(sheet, colIndex, eventItem.EventFile);
        }

        private void ProcessDonationAmountCells(ISheet overviewSheet, int colIndex, string eventFilePath)
        {
            // Locate event donation amount cells (20 cells below event name) in the overview file
            for (int rowIndex = 3; rowIndex <= 22; rowIndex++) // 20 cells below row 3 is row 4 to row 23 (index 3 to 22)
            {
                IRow donationRow = overviewSheet.GetRow(rowIndex);
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
                            int boardMemberId = int.Parse(boardMemberName.Split('.')[0]); // Get identifier

                            // Process event file to find matching donation amounts and board member names
                            ProcessEventFileDetails(eventFilePath, overviewSheet, colIndex, boardMemberId, rowIndex);
                        }
                    }
                }
            }
        }

        private void ProcessEventFileDetails(string eventFilePath, ISheet overviewSheet, int colIndex, int boardMemberId, int overviewRowIndex)
        {
            try
            {
                using (var eventFileStream = new FileStream(eventFilePath, FileMode.Open, FileAccess.Read))
                {
                    // Check for the sheet '贊助記錄總表'
                    IWorkbook eventWorkbook = new XSSFWorkbook(eventFileStream);
                    ISheet eventSheet = eventWorkbook.GetSheet("贊助記錄總表");
                    if (eventSheet == null)
                    {
                        MessageBox.Show("Sheet '贊助記錄總表' not found in event file.");
                        return;
                    }

                    // Find the column for '節目贊助金額'
                    int donationColIndex = FindDonationColumnIndex(eventSheet);
                    if (donationColIndex == -1)
                    {
                        MessageBox.Show("Column '節目贊助金額' not found in event file.");
                        return;
                    }

                    // Process donation amount cells in the event file
                    for (int rowIndex = 5; rowIndex <= 24; rowIndex++) // Row 6 to row 25 (index 5 to 24)
                    {
                        IRow eventRow = eventSheet.GetRow(rowIndex);
                        if (eventRow != null)
                        {
                            ICell eventDonationCell = eventRow.GetCell(donationColIndex);
                            if (eventDonationCell != null && !string.IsNullOrEmpty(eventDonationCell.ToString()))
                            {
                                // Construct the linkage formula
                                string eventFileName = Path.GetFileName(eventFilePath);
                                string eventFileDirectory = Path.GetDirectoryName(eventFilePath);
                                string cellAddress = eventDonationCell.Address.ToString();
                                string sheetName = eventSheet.SheetName;
                                string linkageFormula = $"'{eventFileDirectory}\\[{eventFileName}]{sheetName}'!{cellAddress}";

                                // Look for board member name in column A
                                ICell eventBoardMemberCell = eventRow.GetCell(0); // Column A is index 0
                                if (eventBoardMemberCell != null)
                                {
                                    string eventBoardMemberName = eventBoardMemberCell.ToString();
                                    int eventBoardMemberId = int.Parse(eventBoardMemberName.Split('.')[0]); // Get identifier

                                    // Check if board member identifiers match
                                    if (eventBoardMemberId == boardMemberId)
                                    {
                                        // Paste the linkage formula to the overview file
                                        IRow overviewRow = overviewSheet.GetRow(overviewRowIndex);
                                        if (overviewRow != null)
                                        {
                                            ICell overviewDonationCell = overviewRow.GetCell(colIndex);
                                            if (overviewDonationCell == null)
                                            {
                                                overviewDonationCell = overviewRow.CreateCell(colIndex);
                                            }
                                            overviewDonationCell.SetCellFormula(linkageFormula);
                                        }
                                    }
                                }
                            }
                        }
                    }

                    // Use the SUM function to calculate the total amount of overviewDonationCells from row 4 to row 23 and place the sum in row 24
                    IRow totalRow = overviewSheet.GetRow(23) ?? overviewSheet.CreateRow(23); // Row 24 (index 23)
                    ICell totalCell = totalRow.GetCell(colIndex) ?? totalRow.CreateCell(colIndex);
                    totalCell.SetCellFormula($"SUM({GetCellAddress(3, colIndex)}:{GetCellAddress(22, colIndex)})");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"An error occurred while processing the event file: {ex.Message}");
            }
        }

        private string GetCellAddress(int rowIndex, int colIndex)
        {
            return $"{(char)('A' + colIndex)}{rowIndex + 1}";
        }

        private int FindDonationColumnIndex(ISheet eventSheet)
        {
            IRow headerRow = eventSheet.GetRow(1); // Row 2 is the second row (index 1)
            for (int col = 0; col < headerRow.LastCellNum; col++)
            {
                ICell cell = headerRow.GetCell(col);
                if (cell != null && cell.ToString().Replace("\r\n", string.Empty).Replace("\n", string.Empty) == "節目贊助金額")
                {
                    return col;
                }
            }
            return -1;
        }
    }

    public class Event
    {
        public string EventFile { get; set; }
    }

    public class Overview
    {
        public string OverviewFilePath { get; set; }
    }

    public class OverviewData
    {
        public Overview Overview { get; set; }
        public List<Event> Events { get; set; }
    }
}
