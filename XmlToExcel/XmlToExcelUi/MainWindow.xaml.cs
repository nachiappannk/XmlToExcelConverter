using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using XmlToExcel;

namespace XmlToExcelUi
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }

        private void UIElement_OnDrop(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                string[] files = (string[])e.Data.GetData(DataFormats.FileDrop);
                SetProcessingState();
                try
                {
                    OnDropActual(files);
                    SetInitialState();
                }
                catch (Exception exception)
                {
                    errorMessage.Text = exception.ToString();
                    GetErrorState();
                }
            }
        }

        private void GetErrorState()
        {
            mainGrid.Visibility = Visibility.Collapsed;
            processingGrid.Visibility = Visibility.Collapsed;
            errorGrid.Visibility = Visibility.Visible;
        }

        private void SetInitialState()
        {
            mainGrid.Visibility = Visibility.Visible;
            processingGrid.Visibility = Visibility.Collapsed;
            errorGrid.Visibility = Visibility.Collapsed;
        }

        private void SetProcessingState()
        {
            processingGrid.Visibility = Visibility.Visible;
            mainGrid.Visibility = Visibility.Collapsed;
            errorGrid.Visibility = Visibility.Collapsed;
        }

        private void OnDropActual(string[] files)
        {
            if (files?.Length != 1) return;
            var inputFileName = files[0];
            if (!System.IO.File.Exists(inputFileName)) return;
            var baseFileName = System.IO.Path.GetFileNameWithoutExtension(inputFileName);
            var initialDirectory = System.IO.Path.GetDirectoryName(inputFileName);
            var outputFile = GetOutputFileName(initialDirectory, baseFileName);
            if (File.Exists(outputFile))
                File.Delete(outputFile);
            if (string.IsNullOrEmpty(outputFile)) return;
            XmlToExcelConverter.ConvertXmlToExcel(inputFileName, outputFile);
        }

        private string GetOutputFileName(string initialDirectory, string baseFileName)
        {
            Microsoft.Win32.SaveFileDialog dlg = new Microsoft.Win32.SaveFileDialog();
            dlg.FileName = baseFileName; // Default file name
            dlg.DefaultExt = ".xlsx"; // Default file extension
            dlg.InitialDirectory = initialDirectory;
            dlg.Filter = "Excel documents (.xlsx)|*.xlsx"; // Filter files by extension
            

            // Show save file dialog box
            Nullable<bool> result = dlg.ShowDialog();

            // Process save file dialog box results
            if (result == true)
            {
                // Save document
                return dlg.FileName;
            }
            return string.Empty;
        }

        private void ButtonBase_OnClick(object sender, RoutedEventArgs e)
        {
            errorMessage.Text = "";
            SetInitialState();
        }
    }
}
