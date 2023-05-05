using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Runtime.InteropServices;

//Microsoft Excel 16 object in references-> COM tab
using Excel = Microsoft.Office.Interop.Excel;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;
using Microsoft.Office.Interop.Excel;
using System.Xml.Linq;

namespace WF_Excel_Read_Write_04_05_2023
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
           

        }
        private string FileName = @"C:\Users\Fishman_1\Documents\data.xlsx";

        private void Form1_Load(object sender, EventArgs e)
        {

        }
        // не работает нужно указать путь
        private void buttSave_Click(object sender, EventArgs e)
        {
           
        }
        private void btnWrite_Click(object sender, EventArgs e)
        {

            // Set cursor as hourglass
            Cursor.Current = Cursors.WaitCursor;
            Excel.Application xlApp = new Excel.Application();
            Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(FileName);
            Excel._Worksheet xlWorksheet = xlWorkbook.Sheets[1];
            Excel.Range xlRange = xlWorksheet.UsedRange;

            xlWorksheet.Cells[1, 1] = txtWrite.Text;
            xlApp.Visible = false;
            xlApp.UserControl = false;
            xlWorkbook.Save();

            //cleanup
            GC.Collect();
            GC.WaitForPendingFinalizers();

            //release com objects to fully kill excel process from running in the background
            Marshal.ReleaseComObject(xlRange);
            Marshal.ReleaseComObject(xlWorksheet);

            //close and release
            xlWorkbook.Close();
            Marshal.ReleaseComObject(xlWorkbook);

            //quit and release
            xlApp.Quit();
            Marshal.ReleaseComObject(xlApp);

            // Set cursor as default arrow
            Cursor.Current = Cursors.Default;
        }

        private void btnRead_Click(object sender, EventArgs e)
        {
            // Set cursor as hourglass
            Cursor.Current = Cursors.WaitCursor;

            Excel.Application xlApp = new Excel.Application();
            Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(FileName);
            Excel._Worksheet xlWorksheet = xlWorkbook.Sheets[1];
            Excel.Range xlRange = xlWorksheet.UsedRange;

            if (xlRange.Cells[1, 1] != null && xlRange.Cells[1, 1].Value2 != null)
            {
                txtRead.Text = xlRange.Cells[1, 1].Value2.ToString();
            }

            //cleanup
            GC.Collect();
            GC.WaitForPendingFinalizers();

            //release com objects to fully kill excel process from running in the background
            Marshal.ReleaseComObject(xlRange);
            Marshal.ReleaseComObject(xlWorksheet);

            //close and release
            xlWorkbook.Close();
            Marshal.ReleaseComObject(xlWorkbook);

            //quit and release
            xlApp.Quit();
            Marshal.ReleaseComObject(xlApp);

            // Set cursor as default arrow
            Cursor.Current = Cursors.Default;
        }



        private void saveFileDialog1_FileOk(object sender, CancelEventArgs e)
        {
           
        }
    }
}
