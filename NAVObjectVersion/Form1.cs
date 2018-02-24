using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using Word = Microsoft.Office.Interop.Word;

namespace NAVObjectVersion
{
    public partial class mainForm : Form
    {
        public mainForm()
        {
            InitializeComponent();

            /*
            TODO:
            Error Checking
            Beautify
            Check for wrong clipboard
            check for empty clipboard
            try to skip having a local copy of word document
            user settings for template?
            */
        }

        private void b_LoadClipboard_Click(object sender, EventArgs e)
        {
            b_PasteClipBoard.Enabled = false;
            this.Cursor = Cursors.WaitCursor;

            Excel.Application xlApp = new Excel.Application();

            if (xlApp == null)
            {
                MessageBox.Show("Excel is not properly installed!!");
                return;
            }

            xlApp.Visible = false;
            Excel.Workbook xlWorkBook;
            Excel.Worksheet xlWorkSheet;
            object misValue = System.Reflection.Missing.Value;
            xlWorkBook = xlApp.Workbooks.Add(misValue);
            xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);

            //Fix WorkSheet Font & Font Size
            xlApp.StandardFont = "Arial";
            xlApp.StandardFontSize = 9.5;

            Excel.Range xlCells = xlWorkBook.Worksheets[1].Cells;
            xlCells.NumberFormat = "@";

            int rowID = 1;
            int colID = 1;
            int totalxlRows = 0;
            int totalxlCols = 0;

            using (StringReader reader = new StringReader(Clipboard.GetText()))
            {
                string line = string.Empty;
                do
                {
                    line = reader.ReadLine();
                    if (line != null)
                    {
                        String[] fieldArr = line.Split('\t');
                        foreach (string field in fieldArr)
                        {
                            xlWorkSheet.Cells[rowID, colID] = field;
                            colID++;                            
                        }
                        if (rowID == 1)
                        {
                            totalxlCols = colID - 1;
                        }
                        rowID++;
                        colID = 1;
                    }
                } while ((line != null));
                totalxlRows = rowID - 1;
            }

            //Fix Data
            for (int currCol = 1; currCol <= totalxlCols; currCol++)
            {
                for (int currRow = 1; currRow <= totalxlRows; currRow++)
                {
                    //Fix Object Type
                    if (xlWorkSheet.Cells[1, currCol].Value.ToString() == "Type")
                    {
                        xlWorkSheet.Cells[currRow, currCol] = FormatType(xlWorkSheet.Cells[currRow, currCol].Value.ToString());
                    }

                    //Fix Date
                    if (xlWorkSheet.Cells[1, currCol].Value.ToString() == "Date")
                    {
                        xlWorkSheet.Cells[currRow, currCol] = FormatDate(xlWorkSheet.Cells[currRow, currCol].Value.ToString());
                    }
                }

                if (xlWorkSheet.Cells[1, currCol].Value.ToString() == "Version List")
                {
                    xlWorkSheet.Cells[1, currCol] = "Version";
                }
            }

            //Copy Data To Word
            this.Cursor = Cursors.Default;

            // Create an instance of the Open File Dialog Box
            var openFileDialog1 = new OpenFileDialog();

            // Set filter options and filter index
            openFileDialog1.Filter = "Word Documents (.docx)|*.docx|All files (*.*)|*.*";
            openFileDialog1.FilterIndex = 1;
            openFileDialog1.Multiselect = false;

            // Call the ShowDialog method to show the dialog box.
            openFileDialog1.ShowDialog();
            //txtDocument.Text = openFileDialog1.FileName;

            bool showWord = true;

            Word.Application WordApp = new Word.Application();
            WordApp.Visible = showWord;

            object miss = System.Reflection.Missing.Value;
            object path = openFileDialog1.FileName;

            Random random = new Random();
            int randomNumber = random.Next(0, 100);
            object path2 = path + randomNumber.ToString() + ".docx";
            File.Copy(path.ToString(), path2.ToString(), true);
            object readOnly = false;
            object SaveChanges = true;
            object visible = true;
            
            this.Cursor = Cursors.WaitCursor;

            Word.Document docs = WordApp.Documents.Open(ref path2, ref miss, ref readOnly,
                                           ref miss, ref miss, ref miss, ref miss,
                                           ref miss, ref miss, ref miss, ref visible,
                                           ref miss, ref miss, ref miss, ref miss,
                                           ref miss);


            //Get First Word Table
            Word.Table wordTable = docs.Tables[1];
            Word.Row wordHeadRow = wordTable.Rows[1];
            //Create Extra Lines to fit Excel Data
            for (int i=2; i < totalxlRows; i++) //Template already has 1 line
            {
                wordTable.Rows.Add();
            }

            foreach (Word.Cell c in wordHeadRow.Cells)
            {               
                for (int currxlCol = 1; currxlCol <= totalxlCols; currxlCol++)
                {
                    Console.WriteLine("Excel: " + xlWorkSheet.Cells[1, currxlCol].Value.ToString() + " <=> Word: " + c.Range.Text.ToString());
                    string xlString = xlWorkSheet.Cells[1, currxlCol].Value.ToString();
                    if (xlString.IndexOf(" ") > -1)
                    {
                        xlString = xlString.Substring(0, xlString.IndexOf(" "));
                    }
                    string wordString = c.Range.Text.ToString();
                    Console.WriteLine("Excel: " + xlString + " <=> Word: " + wordString);
                    //Populate Object Type
                    if (xlWorkSheet.Cells[1, currxlCol].Value.ToString().Trim() == "Type" && c.Range.Text.ToString().IndexOf(xlString) > -1)
                    {
                        for (int currxlRow = 2; currxlRow <= totalxlRows; currxlRow++)
                        {
                            wordTable.Cell(currxlRow, currxlCol).Range.Text = xlWorkSheet.Cells[currxlRow, currxlCol].Value.ToString();
                        }
                    }
                    //Populate Object ID
                    if (xlWorkSheet.Cells[1, currxlCol].Value.ToString().Trim() == "ID" && c.Range.Text.ToString().IndexOf(xlString) > -1)
                    {
                        for (int currxlRow = 2; currxlRow <= totalxlRows; currxlRow++)
                        {
                            wordTable.Cell(currxlRow, currxlCol).Range.Text = xlWorkSheet.Cells[currxlRow, currxlCol].Value.ToString();
                        }
                    }

                    //Populate Name
                    if (xlWorkSheet.Cells[1, currxlCol].Value.ToString().Trim() == "Name" && c.Range.Text.ToString().IndexOf(xlString) > -1)
                    {
                        for (int currxlRow = 2; currxlRow <= totalxlRows; currxlRow++)
                        {
                            wordTable.Cell(currxlRow, currxlCol).Range.Text = xlWorkSheet.Cells[currxlRow, currxlCol].Value.ToString();
                        }
                    }

                    //Populate Modified
                    if (xlWorkSheet.Cells[1, currxlCol].Value.ToString().Trim() == "Modified" && c.Range.Text.ToString().IndexOf(xlString) > -1)
                    {
                        for (int currxlRow = 2; currxlRow <= totalxlRows; currxlRow++)
                        {
                            wordTable.Cell(currxlRow, currxlCol).Range.Text = xlWorkSheet.Cells[currxlRow, currxlCol].Value.ToString();
                        }
                    }

                    //Populate Version
                    if (xlWorkSheet.Cells[1, currxlCol].Value.ToString().Trim() == "Version" && c.Range.Text.ToString().IndexOf(xlString) > -1)
                    {
                        for (int currxlRow = 2; currxlRow <= totalxlRows; currxlRow++)
                        {
                            wordTable.Cell(currxlRow, currxlCol).Range.Text = xlWorkSheet.Cells[currxlRow, currxlCol].Value.ToString();
                        }
                    }

                    //Populate Date
                    if (xlWorkSheet.Cells[1, currxlCol].Value.ToString().Trim() == "Date" && c.Range.Text.ToString().IndexOf(xlString) > -1)
                    {
                        for (int currxlRow = 2; currxlRow <= totalxlRows; currxlRow++)
                        {
                            wordTable.Cell(currxlRow, currxlCol).Range.Text = xlWorkSheet.Cells[currxlRow, currxlCol].Value.ToString();
                        }
                    }

                    //Populate Time
                    if (xlWorkSheet.Cells[1, currxlCol].Value.ToString().Trim() == "Time" && c.Range.Text.ToString().IndexOf(xlString) > -1)
                    {
                        for (int currxlRow = 2; currxlRow <= totalxlRows; currxlRow++)
                        {
                            wordTable.Cell(currxlRow, currxlCol).Range.Text = xlWorkSheet.Cells[currxlRow, currxlCol].Value.ToString();
                        }
                    }
                }
            }


            //xlWorkBook.SaveAs(@"C:\Temp\csharp-Excel.xls", Excel.XlFileFormat.xlWorkbookNormal, misValue, misValue, misValue, misValue, Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue);
            xlWorkBook.Close(false, misValue, misValue);

            //Clear Excel Variables
            xlApp.Application.Quit();
            xlApp.Quit();

            Marshal.ReleaseComObject(xlCells);
            Marshal.ReleaseComObject(xlWorkSheet);
            Marshal.ReleaseComObject(xlWorkBook);
            Marshal.ReleaseComObject(xlApp);

            xlCells = null;
            xlWorkSheet = null;
            xlWorkBook = null;
            xlApp = null;


            //Clear Word variables
            if (!showWord)
            {
                docs.Close(SaveChanges);
                WordApp.Application.Quit(SaveChanges);
                WordApp.Quit();
            }

            Marshal.ReleaseComObject(wordHeadRow);
            Marshal.ReleaseComObject(wordTable);
            Marshal.ReleaseComObject(docs);
            Marshal.ReleaseComObject(WordApp);

            wordHeadRow = null;
            wordTable = null;
            docs = null;
            WordApp = null;

            GC.Collect();
            
            MessageBox.Show("Done");
            this.Cursor = Cursors.Default;
            b_PasteClipBoard.Enabled = true;
        }

        private string FormatType(string objTypeNum)
        {
            switch(objTypeNum)
            {
                case "Type": return "Type";
                case "1": return "Table";
                case "2": return "Form";
                case "3": return "Report";
                case "4": return "Dataport";
                case "5": return "Codeunit";
                case "6": return "XMLport";
                case "7": return "MenuSuite";
                case "8": return "Page";
                case "9": return "Query";
                default: return "";
            }
        }

        private string FormatDate(string oldDate)
        {
            if (oldDate == "Date")
            {
                return "Date";
            }
            else
            {
                return DateTime.ParseExact(oldDate, "dd/MM/yy", CultureInfo.InvariantCulture).ToString("dd/MM/yy");
            }
        }

    }
}
