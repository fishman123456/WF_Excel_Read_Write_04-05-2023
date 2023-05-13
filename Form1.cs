using System;
using System.ComponentModel;
using System.Windows.Forms;
using System.Runtime.InteropServices;
//Microsoft Excel 16 object in references-> COM tab
using Excel = Microsoft.Office.Interop.Excel;
using Microsoft.Office.Interop.Excel;
using System.IO;
using System.Collections.Generic;
using OfficeOpenXml;

namespace WF_Excel_Read_Write_04_05_2023
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }
        //private string FileName = @"C:\Users\Fishman_1\Documents\data.xlsx";
        public string filename2 = "";
        private void Form1_Load(object sender, EventArgs e)
        {

        }
        //// не работает нужно указать путь
        // работает только нужно создать файл
        private void buttSave_Click(object sender, EventArgs e)
        {
            // нужен NuGet EPPlus
            // заработало 14-05-2023 00:07 
            SaveFileDialog saveFile = new SaveFileDialog();
            {
                saveFile.Filter = "(*.xlsx)|*.xlsx|Все файлы (*.*)|*.*\"\"\r\n";
                saveFile.Title = "Сохранить";
            };
            
            ///
            // третий вариант
            //https://stackoverflow.com/questions/64824327/i-am-getting-an-error-while-exporting-to-excel
            //var saveFileDialog = new SaveFileDialog();
            //saveFileDialog.FileName = "";
            //saveFileDialog.DefaultExt = ".xls";
            if (saveFile.ShowDialog() == DialogResult.OK)
            {
                filename2 = saveFile.FileName;
                ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;

                using (ExcelPackage pck = new ExcelPackage(new FileInfo(saveFile.FileName)))
                {
                    try
                    {
                        ExcelWorksheet ws = pck.Workbook.Worksheets.Add("Первый");
                        ws.Cells["A1"].Value = "1";
                        pck.Save();
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.ToString());
                        //throw;
                    }
                   
                }
            }
            #region
            // второй вариант
            //using (ExcelPackage excel = new ExcelPackage())
            //{
            //    excel.Workbook.Worksheets.Add("Worksheet1");

            //    var headerRow = new List<string[]>()
            //        {
            //            new string[] { "id", "Name"}
            //        };

            //    // Determine the header range (e.g. A1:D1)
            //    string headerRange = "A1:" + Char.ConvertFromUtf32(headerRow[0].Length + 64) + "1";

            //    // Target a worksheet
            //    var worksheet = excel.Workbook.Worksheets["Worksheet1"];

            //    // Popular header row data
            //    worksheet.Cells[headerRange].LoadFromArrays(headerRow);
            //    FileInfo excelFile = new FileInfo(saveFile.FileName);
            //    //FileInfo excelFile = new FileInfo(@"C:\dddd.xlsx");
            //    excel.SaveAs(excelFile);
            //}

            //// взято с youtube
            ////https://habr.com/ru/sandbox/122135/
            //var path = Path.GetDirectoryName(saveFile.FileName);
            //var wb = new Workbook();
            //var sh = wb.Worksheets.Add("denu");

            // придумки дениса через filestream
            //FileStream fileStream = new FileStream(saveFile.FileName,
            //    FileMode.Create, FileAccess.Write);
            //filename2 = saveFile.FileName;
            //StreamWriter writer = new StreamWriter(fileStream, Encoding.UTF8);
            ////writer.WriteLine("");

            #endregion
        }
        // обработчик нажатия на клавишу записать в ексел
        private void btnWrite_Click(object sender, EventArgs e)
        {
            try
            {
                Cursor.Current = Cursors.WaitCursor;
            Excel.Application xlApp = new Excel.Application();
            Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(filename2);
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
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString(),"Скорее создайте файл");
                //throw;
            }
            // Set cursor as hourglass
            

            
        }
        // обработчик нажатия на клавишу считать из ексел
        private void btnRead_Click(object sender, EventArgs e)
        {
            try
            {
                // Set cursor as hourglass
                Cursor.Current = Cursors.WaitCursor;

                Excel.Application xlApp = new Excel.Application();
                Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(filename2);
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
            catch (Exception ex)
            {

                MessageBox.Show(ex.ToString(),"Скорее создайте файл");
                //throw;
            }
           
        }



        private void saveFileDialog1_FileOk(object sender, CancelEventArgs e)
        {
           
        }

        private void txtWrite_TextChanged(object sender, EventArgs e)
        {

        }
    }
}
