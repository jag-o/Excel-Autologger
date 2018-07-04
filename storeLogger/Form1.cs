using Excel = Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Windows.Forms;

namespace storeLogger
{
    public partial class Form1 : Form
    {
        string excelDocument; // Global string needed for document.
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
            // A huge block of if statements. This is to be improved in the future.
            // If the cell is empty, set it to undefined.
            // Might be improved in the future with user control over this feature.
            if (Element.Text == "")
            {
                Element.Text = "Undefined";
            }
            if (Quantity.Text == "")
            {
                Quantity.Text = "Undefined";
            }
            if (Hazcard.Text == "")
            {
                Hazcard.Text = "Undefined";
            }
            if (Location.Text == "")
            {
                Location.Text = "Undefined";
            }
            if (boughtIn.Text == "")
            {
                boughtIn.Text = "Undefined";
            }
            if (usedBy.Text == "")
            {
                usedBy.Text = "Undefined";
            }
            if (Stock.Text == "")
            {
                Stock.Text = "Undefined";
            }
            if (Company.Text == "")
            {
                Company.Text = "Undefined";
            }
            if (Comment.Text == "")
            {
                Comment.Text = "Undefined";
            }
            // Again, a huge block of code. Could be optimised in the future with
            // potentially a for-loop, and non-specific strings.
            UpdateExcel((Convert.ToInt32(rowCount.Value)), 2, Element.Text);
            UpdateExcel((Convert.ToInt32(rowCount.Value)), 3, Quantity.Text);
            UpdateExcel((Convert.ToInt32(rowCount.Value)), 4, Hazcard.Text);
            UpdateExcel((Convert.ToInt32(rowCount.Value)), 5, Location.Text);
            UpdateExcel((Convert.ToInt32(rowCount.Value)), 6, boughtIn.Text);
            UpdateExcel((Convert.ToInt32(rowCount.Value)), 7, usedBy.Text);
            UpdateExcel((Convert.ToInt32(rowCount.Value)), 8, Stock.Text);
            UpdateExcel((Convert.ToInt32(rowCount.Value)), 9, Company.Text);
            UpdateExcel((Convert.ToInt32(rowCount.Value)), 10, Comment.Text);
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
            List<string> list = PullWorksheets();
            dropDown.Items.Clear();
            foreach (string name in list)
            {
                dropDown.Items.Add(name);
            }
        }
        private void RowCount_ValueChanged(object sender, EventArgs e)
        {
            // Heavily unoptimised code, takes a while to load up.
            // As you change the row count, the text boxes auto fill with cell information.
            Element.Text = DataString((Convert.ToInt32(rowCount.Value)), 2);
            Quantity.Text = DataString((Convert.ToInt32(rowCount.Value)), 3);
            Hazcard.Text = DataString((Convert.ToInt32(rowCount.Value)), 4);
            Location.Text = DataString((Convert.ToInt32(rowCount.Value)), 5);
            boughtIn.Text = DataString((Convert.ToInt32(rowCount.Value)), 6);
            usedBy.Text = DataString((Convert.ToInt32(rowCount.Value)), 7);
            Stock.Text = DataString((Convert.ToInt32(rowCount.Value)), 8);
            Company.Text = DataString((Convert.ToInt32(rowCount.Value)), 9);
            Comment.Text = DataString((Convert.ToInt32(rowCount.Value)), 10);
        }
        public string DataString(int row, int col)
        {
            // This method extracts data from cells using the .ToString() function.
            // However, as this simply hands the instructions off, Excel is extremely
            // slow reading and sending data and causes a bottleneck. Looking for a better
            // solution.
            Excel.Application oXL = null;
            Excel._Workbook oWB = null;
            Excel._Worksheet oSheet = null;
            oXL = new Excel.Application();
            oWB = oXL.Workbooks.Open(excelDocument);
            oSheet = String.IsNullOrEmpty(dropDown.Text) ? (Excel._Worksheet)oWB.ActiveSheet : (Excel._Worksheet)oWB.Worksheets[dropDown.Text];
            string data = null;
            try
            {
                data = oSheet.Cells[row, col].Value.ToString();
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
            return data; // Return the data inside the cell.
        }
        public int RowMinMax()
        {
            // New method written so the end user doesn't go over the amount of cells already used.
            // However, this may be scratched as it could be that the end user wishes to fill new boxes.
            int rowCounter = 0;
            Excel.Application oXL = new Excel.Application();
            Excel._Workbook oWB = oXL.Workbooks.Open(excelDocument);
            Excel._Worksheet oSheet = null;
            oSheet = String.IsNullOrEmpty(dropDown.Text) ? (Excel._Worksheet)oWB.ActiveSheet : (Excel._Worksheet)oWB.Worksheets[dropDown.Text];
            Excel.Range last = oSheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell, Type.Missing); // Finds the last used cell.
            Excel.Range range = oSheet.get_Range("A1", last);
            try
            {
                rowCounter = last.Row;
            }
            catch (Exception ex)
            {
                Debug.Write(ex);
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
            oXL = new Excel.Application();
            oWB = oXL.Workbooks.Open(excelDocument);
            oSH = oWB.Worksheets;
            // This foreach loop determines each name in the specified workbook, oWB.
            foreach (Excel.Worksheet worksheet in oWB.Worksheets)
            { 
                str.Add(worksheet.Name);
            }
            if (oWB != null)
            {
                oWB.Close(); // Close down document, to avoid deadlock or R/W issues.
            }
            return str; // Return the List<> of worksheets.
        }
        private void DropDown_SelectedIndexChanged(object sender, EventArgs e)
        {
            // This two line code is run whenever the sheet of the workbook is changed,
            // so the min-max value in RowMinMax can be determined.
            rowCount.Maximum = RowMinMax();
            rowCount.Minimum = 1;
        }
    }
}
