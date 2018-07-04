using Excel = Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.IO;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;

namespace storeLogger
{
    public partial class Form1 : Form
    {
        string excelDocument;
        public Form1()
        {
            InitializeComponent();
        }

        private void UpdateExcel(int row, int col, string data)
        {
            Excel.Application oXL = null;
            Excel._Workbook oWB = null;
            Excel._Worksheet oSheet = null;

            try
            {
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
                    oWB.Close();
                }
            }
        }
        private void button1_Click(object sender, EventArgs e)
        {
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

        private void openFileDialog1_FileOk(object sender, CancelEventArgs e)
        {
        }

        public void button2_Click(object sender, EventArgs e)
        {
            openFileDialog1.ShowDialog();
            excelDocument = openFileDialog1.FileName;
            label1.Text = label1.Text + " " + excelDocument;
            List<string> list = pullWorksheets();
            foreach (string name in list)
            {
                dropDown.Items.Add(name);
            }

        }

        private void rowCount_ValueChanged(object sender, EventArgs e)
        {
            Element.Text = dataString((Convert.ToInt32(rowCount.Value)), 2);
            Quantity.Text = dataString((Convert.ToInt32(rowCount.Value)), 3);
            Hazcard.Text = dataString((Convert.ToInt32(rowCount.Value)), 4);
            Location.Text = dataString((Convert.ToInt32(rowCount.Value)), 5);
            boughtIn.Text = dataString((Convert.ToInt32(rowCount.Value)), 6);
            usedBy.Text = dataString((Convert.ToInt32(rowCount.Value)), 7);
            Stock.Text = dataString((Convert.ToInt32(rowCount.Value)), 8);
            Company.Text = dataString((Convert.ToInt32(rowCount.Value)), 9);
            Comment.Text = dataString((Convert.ToInt32(rowCount.Value)), 10);

        }

        public string dataString(int row, int col)
        {
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

            } catch(Exception ex)
            {
                Debug.WriteLine(ex);
            }
            return data;
        }
        public List<string> pullWorksheets()
        {
            List<string> str = new List<string>();
            Excel.Application oXL = null;
            Excel._Workbook oWB = null;
            Excel.Sheets oSH = null;
            oXL = new Excel.Application();
            oWB = oXL.Workbooks.Open(excelDocument);
            oSH = oWB.Worksheets;
            foreach (Excel.Worksheet worksheet in oWB.Worksheets)
            {
                str.Add(worksheet.Name);
            }
            return str;
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }
    }
}
