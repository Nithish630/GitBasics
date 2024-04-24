using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
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

namespace Table
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

        #region ClickEvent
        private void Button_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                //  template document
                using (FileStream fileStreamPath = new FileStream(@"C:\Users\nithi\OneDrive\Desktop\DaikinTools\DaikinTools-Desktop\src\McQuay.McQuayTools.Output.Managers\JobReports\QuoteWorksheetAllItems.dotx", FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                {
                    WordDocument document = new WordDocument(fileStreamPath, FormatType.Docx);


                    List<string[]> tableData = new List<string[]>
            {
                new string[] { "1", "20", "hello", "$50.00" },
                new string[] { "2", "50", "nike", "$100.00" }
            };



                    PopulateTableWithBookmarkData(document, "TUO1", tableData);

                    
                    using (FileStream outFileStream = new FileStream(@"C:\Users\nithi\OneDrive\Documents\Custom Office Templates\OutputTest.docx", FileMode.Create))
                    {
                        document.Save(outFileStream, FormatType.Docx);
                    }

                    document.Close();

                    MessageBox.Show("Data inserted successfully.");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error: " + ex.Message);
            }
        }
        #endregion

        #region Logic
        private void PopulateTableWithBookmarkData(WordDocument document, string bookmarkName, List<string[]> tableData)
        {
            Bookmark bookmark = document.Bookmarks.FindByName(bookmarkName);
            if (bookmark != null && bookmark.BookmarkStart.OwnerParagraph.IsInCell)
            {
                WTableRow templateRow = ((WTableCell)bookmark.BookmarkStart.OwnerParagraph.OwnerTextBody).OwnerRow;
                WTable table = (WTable)templateRow.Owner;

                int rowIndex = table.Rows.IndexOf(templateRow);

                
                table.Rows.RemoveAt(rowIndex);

                foreach (string[] rowData in tableData)
                {
                    WTableRow newRow = templateRow.Clone();

                    for (int i = 0; i < rowData.Length; i++)
                    {
                        if (i < newRow.Cells.Count)
                        {
                            newRow.Cells[i].Paragraphs[0].Text = rowData[i];
                        }
                    }

                    table.Rows.Insert(rowIndex, newRow);
                    rowIndex++;
                }

               
                WSection section = new WSection(document);
                document.Sections.Add(section);

               
                WTable clonedTable = table.Clone();
                section.Tables.Add(clonedTable);

                int rowIndexClone = table.Rows.IndexOf(templateRow) + tableData.Count;


                string[] secondSetData1 = { "10", "30", "nitish", "$100" };
                string[] secondSetData2 = { "20", "40", "satish", "$400" };


                WTableRow newRow1 = templateRow.Clone();
                WTableRow newRow2 = templateRow.Clone();

                for (int i = 0; i < secondSetData1.Length; i++)
                {
                    if (i < newRow1.Cells.Count && i < newRow2.Cells.Count)
                    {
                        newRow1.Cells[i].Paragraphs[0].Text = secondSetData1[i];
                        newRow2.Cells[i].Paragraphs[0].Text = secondSetData2[i];
                    }
                }


                clonedTable.Rows.Insert(rowIndexClone + 1, newRow1);
                clonedTable.Rows.Insert(rowIndexClone + 2, newRow2);


                document.LastSection.Tables.Add(clonedTable);
            }
            else
            {
                Console.WriteLine("Bookmark not found or not in a valid location.");
            }
        }
        #endregion


    }
}
        











    

