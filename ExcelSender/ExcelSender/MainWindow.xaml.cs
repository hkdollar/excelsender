using System;
using System.Windows;
using System.Windows.Shapes;
using Excel = Microsoft.Office.Interop.Excel;
using System.IO;
using System.Reflection;
using Newtonsoft.Json;
using System.Net;
using System.Diagnostics;

namespace ExcelSender
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

        private void BtnImportFile_Click(object sender, RoutedEventArgs e)
        {
            Excel.Application xlApp = null;
            Excel.Workbook xlWorkBook = null;
            Excel.Worksheet xlWorkSheet = null;
            Excel.Range range = null;

            object misValue = System.Reflection.Missing.Value;

            try
            {
                string URL = @"http://iotdtg.w7.n3n.io";
                string path = @"/iotdtg/death";
                int port = 9998;

                var spreadsheetLocation = System.IO.Path.Combine(Directory.GetCurrentDirectory(), "Raw Data.xlsx");

                xlApp = new Excel.Application();
                xlWorkBook = xlApp.Workbooks.Open(spreadsheetLocation, 0, true, 5, "", "", true, Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
                //xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);

                foreach (Excel.Worksheet sheet in xlWorkBook.Worksheets)
                {

                    //Console.WriteLine(@"Values for Sheet " + sheet.Index);

                    // get a range to work with
                    range = sheet.get_Range("A1", Missing.Value);
                    // get the end of values to the right (will stop at the first empty cell)
                    range = range.get_End(Excel.XlDirection.xlToRight);
                    // get the end of values toward the bottom, looking in the last column (will stop at first empty cell)
                    range = range.get_End(Excel.XlDirection.xlDown);

                    // get the address of the bottom, right cell
                    string downAddress = range.get_Address(
                        false, false, Excel.XlReferenceStyle.xlA1,
                        Type.Missing, Type.Missing);

                    // Get the range, then values from a1
                    range = sheet.get_Range("A1", downAddress);
                    object[,] values = (object[,])range.Value2;

                    // View the values
                    //Console.Write("\t");
                    //Console.WriteLine();

                    for (int i = 2; i <= values.GetLength(0); i++)
                    //for (int i = 2; i <= 2; i++)
                    {
                        var d = new ExcelData(values[i, 1], values[i, 2], values[i, 3], values[i, 4], values[i, 5], values[i, 6],
                                        values[i, 7], values[i, 8], values[i, 9], values[i, 10], values[i, 11], values[i, 12], values[i, 13]);

                        var httpWebRequest = (HttpWebRequest)WebRequest.Create(URL + ":" + port.ToString() + path);
                        httpWebRequest.ContentType = "application/json";
                        httpWebRequest.Method = "POST";

                        using (var streamWriter = new StreamWriter(httpWebRequest.GetRequestStream()))
                        {
                            string jsonString = JsonConvert.SerializeObject(d);

                            streamWriter.Write(jsonString);
                            streamWriter.Flush();
                            streamWriter.Close();                            
                        }

                        var httpResponse = (HttpWebResponse)httpWebRequest.GetResponse();
                        using (var streamReader = new StreamReader(httpResponse.GetResponseStream()))
                        {
                            var result = streamReader.ReadToEnd();
                            Debug.WriteLine(String.Format("Sent line No.{0}, Result is {1}", i, result));
                            //MessageBox.Show(result);
                        }
                    }

                    MessageBox.Show("Sent total data");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            finally
            {
                if (xlWorkSheet != null)
                {
                    releaseObject(xlWorkSheet);
                }

                if (xlWorkBook != null)
                {
                    xlWorkBook.Close(true, misValue, misValue);
                    releaseObject(xlWorkBook);
                }

                if (xlApp != null)
                {
                    xlApp.Quit();
                    releaseObject(xlApp);
                }
            }

        }

        private void releaseObject(object obj)
        {
            try
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
                obj = null;
            }
            catch (Exception ex)
            {
                obj = null;
                MessageBox.Show("Unable to release the Object " + ex.ToString());
            }
            finally
            {
                GC.Collect();
            }
        }
    }
}
