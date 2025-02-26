﻿using System;
using System.Collections.Generic;
using System.IO;
using System.Windows;
using Microsoft.Win32;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System.Collections.ObjectModel;

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

        private void SynchronizeButton_Click(object sender, RoutedEventArgs e)
        {
            // This method can be used for further processing if needed
        }
    }

    public class Event
    {
        public string Name { get; set; }
        public string EventFile { get; set; }
    }
}
