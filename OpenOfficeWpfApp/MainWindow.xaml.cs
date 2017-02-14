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
using Microsoft.Win32;

namespace OpenOfficeWpfApp
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

        private void button_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            if (openFileDialog.ShowDialog() == true)
            {
                SelectedFile.Text = openFileDialog.FileName;
                ChangeButton.IsEnabled = true;
            }

        }

        private void ChangeButton_Click(object sender, RoutedEventArgs e)
        {
            var testOfficePackage = new ChangeTestOfficeFileClass();
            try
            {
                testOfficePackage.ChangePackage(SelectedFile.Text);
            }
            catch (Exception exception)
            {
                var message = exception.Message;
                if (exception.Source == "DocumentFormat.OpenXml")
                {
                    message = "Wrong file context, you should specify original text.docx file";
                }
                MessagesTextBox.Text += "\n" + message;
                return;
            }
            MessagesTextBox.Text += "\nFile has been changed successfully";
        }
    }
}
