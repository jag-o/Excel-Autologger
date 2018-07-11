using Excel = Microsoft.Office.Interop.Excel;
using Microsoft.CSharp;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Diagnostics;
using System.Windows.Forms;

namespace storeLogger
{
    public partial class Form1 : Form
    {
        string excelDocument; // Global string needed for document.
        public bool autoFill; // Global logic needed for autofilling of blanks in application.
        Form1 fm;
        public Form1() // Initilises the UI.
        {
            InitializeComponent();
        }
        private void UpdateExcel(int row, int col, string data)
        {
            // UpdateExcel() is the method used to replace cells inside the specified
            // Excel spreadsheet, using the Interop.Excel library.
            Excel.Application oXL = null;
            Excel._Workbook oWB = null;
            Excel._Worksheet oSheet = null;
            try
            {
                // This try block opens up your specified document in the File Chooser,
                // and attempts to write to cells with the specified data given (string data).
                oXL = new Excel.Application();
                oWB = oXL.Workbooks.Open(excelDocument);
                oSheet = String.IsNullOrEmpty(dropDown.Text) ? (Excel._Worksheet)oWB.ActiveSheet : (Excel._Worksheet)oWB.Worksheets[dropDown.Text];
                oSheet.Cells[row, col] = data;
                oWB.Save();
            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex);
            }
            finally
            {
                if (oWB != null)
                {
                    oWB.Close(); // Close down document, to avoid deadlock or R/W issues.
                }
            }
        }
        private void Button1_Click(object sender, EventArgs e)
        {
            // Compacted down from if statements in 0.2.0-beta.
            // This new method creats an array from the boxes, to use a for-loop instead.
            string[] boxArr = new string[9] { s1.Text, s2.Text, s3.Text, s4.Text, s5.Text, s6.Text, s7.Text, s8.Text, s9.Text };
            if (!checkBox1.Checked)
            {
                // Do nothing
            }
            else if (checkBox1.Checked)
            {
                for (int i = 0; i < 9; i++)
                {
                    if (boxArr[i] == "")
                    {
                        boxArr[i] = "Undefined";
                    }
                }
            }
            for (int i = 0; i < 9; i++)
            {
                UpdateExcel((Convert.ToInt32(rowCount.Value)), i, boxArr[i]);
            }
            MessageBox.Show("Done!");
        }
        public void Button2_Click(object sender, EventArgs e)
        {
            // After the file opener button has been clicked, the global
            // string for the file path is set, for later use in methods.
            // It then extracts the sheet names using the PullWorksheets method,
            // and changes the drop box.
            openFileDialog1.ShowDialog();
            excelDocument = openFileDialog1.FileName;
            label1.Text = "Current document open: " + excelDocument;
            // Logic is used to fix a bug where the Excel document would still
            // try to be used, but caused a COM exception at runtime (0.0.12-alpha)
            if (excelDocument == null)
            {
                MessageBox.Show("Please choose a vaild .XLSX file!");
                label1.Text = "Current document open: ";
            }
            else
            {
                List<string> list = PullWorksheets();
                dropDown.Items.Clear();
                foreach (string name in list)
                {
                    dropDown.Items.Add(name);
                }
            }
        }
        private void RowCount_ValueChanged(object sender, EventArgs e)
        {
            QueryData();
        }
        public void QueryData()
        {
            // Heavily unoptimised code, takes a while to load up.
            // As you change the row count, the text boxes auto fill with cell information.
            string[] boxArr = new string[9] { s1.Text, s2.Text, s3.Text, s4.Text, s5.Text, s6.Text, s7.Text, s8.Text, s9.Text };
            for(int i = 0; i < 9; i++)
            {
                boxArr[i] = DataString((Convert.ToInt32(rowCount.Value)), (i + 1));
            }
        }
        public string DataString(int row, int col)
        {
            // This method extracts data from cells using the .ToString() function.
            // However, as this simply hands the instructions off, Excel is extremely
            // slow reading and sending data and causes a bottleneck. Looking for a better
            // solution.
            Excel.Application oXL = new Excel.Application();
            Excel._Workbook oWB = null;
            string data = "";
            try
            {
                // Why in 0.0.12, I forgot to encase this in a try/catch, is beyond me.
                oWB = oXL.Workbooks.Open(excelDocument);
                Excel._Worksheet oSheet = String.IsNullOrEmpty(dropDown.Text) ? (Excel._Worksheet)oWB.ActiveSheet : (Excel._Worksheet)oWB.Worksheets[dropDown.Text];
                data = oSheet.Cells[row, col].Value.ToString();
            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex);
            }
            finally
            {
                if(oWB != null)
                {
                    oWB.Close();
                }
            }
            return data; // Return the data inside the cell.
        }
        public int RowMinMax()
        {
            // New method written so the end user doesn't go over the amount of cells already used.
            // However, this may be scratched as it could be that the end user wishes to fill new boxes.
            int rowCounter = 0;
            Excel.Application oXL = null;
            Excel._Workbook oWB = null;
            Excel._Worksheet oSheet = null;
            try
            {
                // Again, unlike 0.0.12, this should always be in a try/catch.
                oXL = new Excel.Application();
                oWB = oXL.Workbooks.Open(excelDocument);
                oSheet = String.IsNullOrEmpty(dropDown.Text) ? (Excel._Worksheet)oWB.ActiveSheet : (Excel._Worksheet)oWB.Worksheets[dropDown.Text];
                Excel.Range last = oSheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell, Type.Missing); // Finds the last used cell.
                Excel.Range range = oSheet.get_Range("A1", last);
                rowCounter = last.Row;
            }
            catch (Exception ex)
            {
                Debug.Write(ex);
                MessageBox.Show("Please open a valid .xlsx document!");
            }
            finally
            {
                if (oWB != null)
                {
                    oWB.Close(); // Close down document, to avoid deadlock or R/W issues.
                }
            }
            return rowCounter; // Return this max value.
        }
        public List<string> PullWorksheets()
        {
            // This method pulls the worksheet names from the file and stores it
            // in a List<string> datatype, using a foreach loop.
            List<string> str = new List<string>();
            Excel.Application oXL = null;
            Excel._Workbook oWB = null;
            Excel.Sheets oSH = null;
            try
            {
                // Put into try/catch to avoid COM exceptions.
                oXL = new Excel.Application();
                oWB = oXL.Workbooks.Open(excelDocument);
                oSH = oWB.Worksheets;
            }
            catch (Exception ex)
            {
                label1.Text = "Current document open: ";
                Debug.WriteLine(ex);
                MessageBox.Show("Please open a valid .xlsx document!");
            }
            // This foreach loop determines each name in the specified workbook, oWB.
            try
            {
                // This needs to be in a try/catch to avoid errors if the above exception
                // is caught.
                foreach (Excel.Worksheet worksheet in oWB.Worksheets)
                {
                    str.Add(worksheet.Name);
                }
            }
            catch(Exception ex)
            {
                Debug.WriteLine(ex);
            }
            finally
            {
                if (oWB != null)
                {
                    oWB.Close(); // Close down document, to avoid deadlock or R/W issues.
                }
            }
            return str; // Return the List<> of worksheets.
        }
        private void DropDown_SelectedIndexChanged(object sender, EventArgs e)
        {
            // This two line code is run whenever the sheet of the workbook is changed,
            // so the min-max value in RowMinMax can be determined.
            rowCount.Maximum = RowMinMax();
            rowCount.Minimum = 1;
            QueryData(); // Changed in 0.2.0-beta to fix a bug.
        }

        private void button1_Click_1(object sender, EventArgs e)
        {
            TextBox newText = new TextBox();
            newText.Location = new Point(63, (s9.Top + 26));
            Form1 Form1 = new Form1();
            Form1.Controls.Add(newText);
            Form1.Refresh();
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }
    }
}
