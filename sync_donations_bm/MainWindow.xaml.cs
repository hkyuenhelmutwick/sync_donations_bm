using System;
using System.Collections.ObjectModel;
using System.IO;
using System.Windows;
using System.Windows.Controls;
using Microsoft.Win32;

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
            // Further processing if needed
            MessageBox.Show("Synchronization process started.");
        }
    }

    public class Event
    {
        public string Name { get; set; }
        public string EventFile { get; set; }
    }
}
