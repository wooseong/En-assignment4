using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;


namespace schedule
{
    public class FindLecture
    {
        Excel.Application excelApplication;
        Excel.Workbook excelWorkbook;
        Excel._Worksheet excelWorkSheet;
        Excel.Range excelRange;
        int rowCount; //
        int colCount;

        public FindLecture(string directory)
        {
            excelApplication = new Excel.Application(); // Excel 첫번째 워크시트 가져오기
            excelWorkbook = excelApplication.Workbooks.Open(directory);
            excelWorkSheet = excelWorkbook.Sheets[1];
            excelRange = excelWorkSheet.UsedRange;

            rowCount = excelRange.Rows.Count;
            colCount = excelRange.Columns.Count;
        }
        ~FindLecture()
        {

            GC.Collect();
            GC.WaitForPendingFinalizers();
            //release com objects to fully kill excel process from running in the background
            Marshal.ReleaseComObject(excelRange);
            Marshal.ReleaseComObject(excelWorkSheet);

            //close and release
            excelWorkbook.Close();
            Marshal.ReleaseComObject(excelWorkbook);

            //quit and release
            excelApplication.Quit();
            Marshal.ReleaseComObject(excelApplication);
        }

        public void initExcel()
        {
            Console.WriteLine(rowCount);
            Console.WriteLine(colCount);
        }
        public void searchLecture(string nLecture)
        {

            for (int i = 1; i <= rowCount; i++)
            {

                if (excelRange.Cells[i, 2].Value2.ToString() != nLecture)
                {
                    continue;
                }

                for (int j = 1; j <= colCount; j++)
                {
                    //new line
                    //  if (j == 1)
                    //  Console.Write("\r\n");

                    //write the value to the console
                    if (excelRange.Cells[i, j] != null && excelRange.Cells[i, j].Value2 != null)
                        Console.Write(excelRange.Cells[i, j].Value2.ToString() + "\t ");

                }
                Console.WriteLine("\r");
            }

        }
    }
}
