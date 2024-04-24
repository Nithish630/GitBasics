
using DocumentFormat.OpenXml;
using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection.Metadata;
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
using System.IO;
using Syncfusion.DocIO.DLS;
using Syncfusion.DocIO;

namespace DaikinToolsPractice
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
       

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            try
            {
               
                FileStream fileStreamPath = new FileStream(@"C:\Users\nithi\OneDrive\Documents\Custom Office Templates\Assignment1.dotx", FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
                WordDocument document = new WordDocument(fileStreamPath, FormatType.Dotx);

              
                BookmarksNavigator bookmarkNavigator = new BookmarksNavigator(document);
               
                
                bookmarkNavigator.MoveToBookmark("Name");
                bookmarkNavigator.DeleteBookmarkContent(true);
                

                bookmarkNavigator.InsertText("NithishRaju");

                
                bookmarkNavigator.MoveToBookmark("Age");
                bookmarkNavigator.DeleteBookmarkContent(false);
                
                bookmarkNavigator.InsertText("20");
                
                
                MemoryStream stream = new MemoryStream();
               
                document.Save(stream, FormatType.Docx);

                
                stream.Position = 0;

                
                using (FileStream outFileStream = new FileStream(@"C:\Users\nithi\OneDrive\Documents\Custom Office Templates\Output.docx", FileMode.Create))
                {
                    stream.CopyTo(outFileStream);
                }

                
                document.Close();

               
                MessageBox.Show("Text inserted successfully.");
            }
            catch (Exception ex)
            {
                
                MessageBox.Show("Error: " + ex.Message);
            }
        }
    }
}

