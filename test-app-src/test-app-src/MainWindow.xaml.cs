using Microsoft.Win32;
using System;
using System.Windows;
using System.Windows.Controls;

namespace test_app_src
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        #region Private fields
        private string _docxFilePath;
        #endregion
        #region Constructor
        public MainWindow()
        {
            InitializeComponent();
            _docxFilePath = string.Empty;
        }
        #endregion
        #region Handlers
        /// <summary>
        /// "Choose file" button handler
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Open_File_Button_Handler(object sender, EventArgs e)
        {
            try
            {
                OpenFileDialog openFileDialog = new OpenFileDialog()
                {
                    Filter = "Microsoft word file (*.docx,*.doc)|*.docx;*.doc"
                };
                if (openFileDialog.ShowDialog() == true)
                {
                    _docxFilePath = openFileDialog.FileName;
                }
            }
            catch
            {
                MessageBox.Show("Error while reading a file");
            }
        }
        /// <summary>
        /// Handles "Start" button click
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private async void Start_Button_Handler(object sender, EventArgs e)
        {
            Button button = sender as Button;
            
            if (_docxFilePath == string.Empty)
            {
                MessageBox.Show("Choose the file");
                return;
            }
            button.IsEnabled = false;
            // Creating new helper class, which allows us to make changes in file
            WordHelper helper = new WordHelper(_docxFilePath, userInfoTextBox.Text);
            if (await helper.AddQrCodes())
            {
                SaveFileDialog saveFileDialog = new SaveFileDialog()
                {
                    Filter = "Microsoft word file (*.docx,*.doc)|*.docx;*.doc"
                };
                if (saveFileDialog.ShowDialog() == true && helper.Save(saveFileDialog.FileName))
                {
                    MessageBox.Show("File saved successfully!");
                }
                else
                {
                    MessageBox.Show("Error while saving the file");
                }
            }
            else
            {
                MessageBox.Show("Error occured while working with file (probably, no content)");
            }
            button.IsEnabled = true;
        }
    }
    #endregion
}