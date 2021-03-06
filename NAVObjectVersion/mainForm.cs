﻿using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using Word = Microsoft.Office.Interop.Word;

namespace NAVObjectVersion
{
    public partial class MainForm : Form
    {

        List<string> tempFilesCreated = new List<string>();
        int FilesRemaining;

        public MainForm()
        {
            InitializeComponent();

            RefreshUIFromSettings();
        }

        private void b_LoadClipboard_Click(object sender, EventArgs e)
        {
            b_PasteClipBoard.Text = "Preparing copied Data";
            b_PasteClipBoard.Enabled = false;
            this.Cursor = Cursors.WaitCursor;

            if (CheckSettingsAndClipboard())
            {
                //Clipboard is OK, start processing
                if (ProcessClipboard())
                {
                    if (Properties.Settings.Default.UseClipboard)
                    {
                        MessageBox.Show("Object list copied to Clipboard." + "\n" + "You can now paste it anywhere you want! :)", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                    else
                    {
                        MessageBox.Show("Finished! You can now go to Word! :)", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                }
            }

            this.Cursor = Cursors.Default;
            b_PasteClipBoard.Text = "Paste NAV Object list";
            b_PasteClipBoard.Enabled = true;
        }

        private bool CheckSettingsAndClipboard()
        {
            //Check Settings
            string templatePath = Properties.Settings.Default.TemplatePath;
            string templateFileName = Properties.Settings.Default.TemplateFileName;

            if ((templatePath == "") || (templateFileName == ""))
            {
                MessageBox.Show("You haven't selected any template. Please restart your process.", "Template Selection Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                Properties.Settings.Default.TemplatePath = "";
                Properties.Settings.Default.TemplateFileName = "";
                RefreshUIFromSettings();
                return false;
            }
            else
            {
                if (!File.Exists(templatePath))
                {
                    MessageBox.Show("File " + templatePath + " doesn't exist. Select new template", "Template Selection Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    Properties.Settings.Default.TemplatePath = "";
                    Properties.Settings.Default.TemplateFileName = "";
                    RefreshUIFromSettings();
                    return false;
                }
            }

            //Check if Clipboard contains NAV objects list
            bool ClipBoardErrorsExist = false;
            using (StringReader reader = new StringReader(Clipboard.GetText()))
            {
                string line = string.Empty;
                line = reader.ReadLine();
                if (line != null)
                {
                    if (line.Substring(0, 7) != "Type\tID")
                    {
                        ClipBoardErrorsExist = true;
                        MessageBox.Show("Content is not a NAV Object List", "Clipboard Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
                else
                {
                    ClipBoardErrorsExist = true;
                    MessageBox.Show("Clipboard is empty.", "Clipboard Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
            return !ClipBoardErrorsExist;
        }

        private bool ProcessClipboard()
        {
            Excel.Application xlApp = new Excel.Application();
            Word.Application WordApp = new Word.Application();

            if (xlApp == null)
            {
                MessageBox.Show("Excel is not properly installed!!", "Excel COM error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }

            if (WordApp == null)
            {
                MessageBox.Show("Word is not properly installed!!", "Word COM error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }

            object miss = System.Reflection.Missing.Value; //Common for Excel and Word

            b_PasteClipBoard.Text = "Processing Clipboard to Excel";

            xlApp.Visible = false;
            Excel.Workbook xlWorkBook;
            Excel.Worksheet xlWorkSheet;
            xlWorkBook = xlApp.Workbooks.Add(miss);
            xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);

            //Fix WorkSheet Font & Font Size
            xlApp.StandardFont = "Arial";
            xlApp.StandardFontSize = 9.5;


            //Format all Cells are Text
            Excel.Range xlCells = xlWorkBook.Worksheets[1].Cells;
            xlCells.NumberFormat = "@";

            int rowID = 1;
            int colID = 1;
            int totalxlRows = 0;
            int totalxlCols = 0;

            //Copy Clipboard content to Excel Worksheet
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

            //Check that we have imported at least one full row
            if ((totalxlRows == 0) || (totalxlCols == 0))
            {
                MessageBox.Show("No data were copied to Excel", "Excel Processing Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
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
            }


            //Finished importing data to Excel Proceed with Word
            b_PasteClipBoard.Text = "Processing Excel to Word";

            //Copy Data To Word
            this.Cursor = Cursors.Default;
            this.Cursor = Cursors.WaitCursor;

            bool showWord = !Properties.Settings.Default.UseClipboard;
            WordApp.Visible = showWord;
            object visible = true;
            object SaveChanges = true;

            string templatePath = Properties.Settings.Default.TemplatePath;
            string templateFileName = Path.GetFileNameWithoutExtension(Properties.Settings.Default.TemplateFileName);
            string TempDir = @Path.GetDirectoryName(templatePath) + @"\";

            object workingPath = TempDir + templateFileName + "-" + DateTime.Now.ToString("yyyyMMddHHmmss") + ".docx";
            //Copy the template so we don't make any changes in it.
            File.Copy(templatePath.ToString(), workingPath.ToString(), true);

            tempFilesCreated.Add(workingPath.ToString());

            //Open Word Template
            Word.Document docs = WordApp.Documents.Open(ref workingPath, ref miss, ref miss,
                                        ref miss, ref miss, ref miss, ref miss,
                                        ref miss, ref miss, ref miss, ref visible,
                                        ref miss, ref miss, ref miss, ref miss,
                                        ref miss);

            //Get First Word Table
            Word.Table wordTable = docs.Tables[1];

            //Create Extra Lines to fit Excel Data
            for (int i = 2; i < totalxlRows; i++) //Template already has 1 line
            {
                wordTable.Rows.Add();
            }

            //Fill Table with Excel Data
            foreach (Word.Cell c in wordTable.Rows[1].Cells) //Loop through Word Table Columns
            {
                for (int currxlCol = 1; currxlCol <= totalxlCols; currxlCol++) //Loop through Excel Columns
                {
                    string xlString = xlWorkSheet.Cells[1, currxlCol].Value.ToString();
                    if (xlString.IndexOf(" ") > -1)
                    {
                        xlString = xlString.Substring(0, xlString.IndexOf(" "));
                    }
                    string wordString = c.Range.Text.ToString();
                    Console.WriteLine("Excel: " + xlString + " <=> Word: " + wordString);

                    //Populate Object Type
                    if (xlString == "Type" && wordString.IndexOf(xlString) > -1)
                    {
                        for (int currxlRow = 2; currxlRow <= totalxlRows; currxlRow++)
                        {
                            wordTable.Cell(currxlRow, currxlCol).Range.Text = xlWorkSheet.Cells[currxlRow, currxlCol].Value.ToString();
                        }
                    }
                    //Populate Object ID
                    if (xlString == "ID" && wordString.IndexOf(xlString) > -1)
                    {
                        for (int currxlRow = 2; currxlRow <= totalxlRows; currxlRow++)
                        {
                            wordTable.Cell(currxlRow, currxlCol).Range.Text = xlWorkSheet.Cells[currxlRow, currxlCol].Value.ToString();
                        }
                    }

                    //Populate Name
                    if (xlString == "Name" && wordString.IndexOf(xlString) > -1)
                    {
                        for (int currxlRow = 2; currxlRow <= totalxlRows; currxlRow++)
                        {
                            wordTable.Cell(currxlRow, currxlCol).Range.Text = xlWorkSheet.Cells[currxlRow, currxlCol].Value.ToString();
                        }
                    }

                    //Populate Modified
                    if (xlString == "Modified" && wordString.IndexOf(xlString) > -1)
                    {
                        for (int currxlRow = 2; currxlRow <= totalxlRows; currxlRow++)
                        {
                            wordTable.Cell(currxlRow, currxlCol).Range.Text = xlWorkSheet.Cells[currxlRow, currxlCol].Value.ToString();
                        }
                    }

                    //Populate Version
                    if (xlString == "Version" && wordString.IndexOf(xlString) > -1)
                    {
                        for (int currxlRow = 2; currxlRow <= totalxlRows; currxlRow++)
                        {
                            wordTable.Cell(currxlRow, currxlCol).Range.Text = xlWorkSheet.Cells[currxlRow, currxlCol].Value.ToString();
                        }
                    }

                    //Populate Date
                    if (xlString == "Date" && wordString.IndexOf(xlString) > -1)
                    {
                        for (int currxlRow = 2; currxlRow <= totalxlRows; currxlRow++)
                        {
                            wordTable.Cell(currxlRow, currxlCol).Range.Text = xlWorkSheet.Cells[currxlRow, currxlCol].Value.ToString();
                        }
                    }

                    //Populate Time
                    if (xlString == "Time" && wordString.IndexOf(xlString) > -1)
                    {
                        for (int currxlRow = 2; currxlRow <= totalxlRows; currxlRow++)
                        {
                            wordTable.Cell(currxlRow, currxlCol).Range.Text = xlWorkSheet.Cells[currxlRow, currxlCol].Value.ToString();
                        }
                    }
                }
            }

            //Close Excel
            xlWorkBook.Close(false, miss, miss);

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
                //Copy Table to Clipboard
                docs.ActiveWindow.Selection.WholeStory();
                docs.ActiveWindow.Selection.Copy();

                //Close Word
                docs.Close(SaveChanges);
                WordApp.Application.Quit(SaveChanges);
                WordApp.Quit();
            }

            Marshal.ReleaseComObject(wordTable);
            Marshal.ReleaseComObject(docs);
            Marshal.ReleaseComObject(WordApp);

            wordTable = null;
            docs = null;
            WordApp = null;

            GC.Collect();

            return true;
        }

        //Field Formatting Functions
        private string FormatType(string objTypeNum)
        {
            switch (objTypeNum)
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

        private void mainForm_FormClosing(object sender, FormClosingEventArgs e)
        {
            closeForm(true);
        }

        private void closeForm(bool CalledByUser)
        {
            if (CalledByUser)
                FilesRemaining = tempFilesCreated.Count;

            if (tempFilesCreated.Count > 0)
            {
                foreach (string tempFile in tempFilesCreated)
                {
                    FilesRemaining -= 1;
                    try
                    {
                        File.Delete(tempFile);
                        //tempFilesCreated.Remove(tempFile);
                    }
                    catch (Exception ex)
                    {
                        if (FilesRemaining <= 0)
                        {
                            DialogResult dialogResult = MessageBox.Show("You have some NAV version Word Documents open and they cannot be deleted. " + "\n" +
                                                                        "Please close them first and then click Yes in this dialog." + "\n" +
                                                                        "Do you want to retry to correctly close the application now?"
                                                                        , "Delete Temporary Word Files Warning", MessageBoxButtons.YesNo);
                            if (dialogResult == DialogResult.Yes)
                            {
                                closeForm(false);
                            }
                            else if (dialogResult == DialogResult.No)
                            {
                                MessageBox.Show("Some Word Documents created, were not deleted as they are still open.\nPlease delete them manually", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                            }
                        }
                        else
                        {
                            closeForm(false);
                        }
                    }
                }
            }
        }

        private void txt_TemplatePath_DoubleClick(object sender, EventArgs e)
        {
            // Create an instance of the Open File Dialog Box
            var templateFileDialog = new OpenFileDialog();

            // Set filter options and filter index
            templateFileDialog.Title = "Select Table Template  Word File";
            templateFileDialog.Filter = "Word Documents (.docx)|*.docx";
            templateFileDialog.FilterIndex = 1;
            templateFileDialog.Multiselect = false;
            templateFileDialog.ShowDialog();
            string templatePath = templateFileDialog.FileName;
            string templateFileName = Path.GetFileNameWithoutExtension(templateFileDialog.FileName);
            //Check if user selected a template file or cancelled.
            if (templateFileName == null || templateFileName == "")
            {
                MessageBox.Show("You haven't select any template. Please restart your process.", "Template Selection Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            Properties.Settings.Default.TemplatePath = @templatePath;
            Properties.Settings.Default.TemplateFileName = templateFileName;
            Properties.Settings.Default.Save();

            RefreshUIFromSettings();
        }

        private void chb_UseClipboard_CheckedChanged(object sender, EventArgs e)
        {
            Properties.Settings.Default.UseClipboard = chb_UseClipboard.Checked;
            Properties.Settings.Default.Save();

            RefreshUIFromSettings();
        }

        private void RefreshUIFromSettings()
        {
            Properties.Settings.Default.Reload();
            txt_TemplatePath.Text = Properties.Settings.Default.TemplateFileName;
            chb_UseClipboard.Checked = Properties.Settings.Default.UseClipboard;
        }

        private void lbl_github_Click(object sender, EventArgs e)
        {
            Process.Start("https://github.com/ktheod");
        }
    }
}
