using System;
using System.Collections.Generic;
using System.IO;
using System.Windows;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

namespace sync_donations_bm
{
    public partial class MainWindow : Window
    {
        private const string FilePath = @"S:\ITD\Kei\PRDFile_Test3\董事會成員定額紀錄 2024 test.xlsx";

        public MainWindow()
        {
            InitializeComponent();
        }

        private void SynchronizeButton_Click(object sender, RoutedEventArgs e)
        {
            if (!File.Exists(FilePath))
            {
                MessageBox.Show("File not found.");
                return;
            }

            using (var fileStream = new FileStream(FilePath, FileMode.Open, FileAccess.Read))
            {
                IWorkbook workbook = new XSSFWorkbook(fileStream);
                ISheet sheet = workbook.GetSheet("節目贊助");
                if (sheet == null)
                {
                    MessageBox.Show("Sheet '節目贊助' not found.");
                    return;
                }

                var events = new List<Event>();
                IRow row = sheet.GetRow(2); // F3 is the third row (index 2)
                if (row != null)
                {
                    for (int col = 5; col < row.LastCellNum; col++) // F is the sixth column (index 5)
                    {
                        ICell cell = row.GetCell(col);
                        if (cell != null)
                        {
                            string eventName = cell.ToString().Replace("\r\n",string.Empty).Replace("\n",string.Empty);
                            if (!string.IsNullOrEmpty(eventName))
                            {
                                events.Add(new Event { Name = eventName });
                            }
                        }
                    }
                }

                // Process the events as needed
                MessageBox.Show($"Found {events.Count} events.");
            }
        }
    }

    public class Event
    {
        public string Name { get; set; }
    }
}
