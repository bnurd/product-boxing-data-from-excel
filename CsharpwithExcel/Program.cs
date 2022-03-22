using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;

namespace CsharpwithExcel
{
    class Program
    {
        static void Main(string[] args)
        {

            List<Box> boxes = new List<Box>
            {
                new Box
                {
                    height = 30,
                    width = 20,
                    tick = 10
                },
                new Box
                {
                    height = 15,
                    width = 10,
                    tick = 5
                },
                new Box
                {
                    height = 10,
                    width = 10,
                    tick = 10
                }
            };
            // DisplayInExcel(boxes);
            ReadFromExcel(@"c:\exceldb.xlsx");
            Console.Read();
        }

        static void ReadFromExcel(string filename)
        {
            Application application = new Application();
            application.Visible = false;

            Workbook workbook = application.Workbooks.Open(filename, 0, true, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows
               , "\t", false, false, 0, true, 1, 0);

            Worksheet worksheet = (Worksheet)workbook.Worksheets.get_Item(1);

            string str;
            int rCnt;
            int cCnt;
            int rw = 0;
            int cl = 0;

            Range range = worksheet.UsedRange;
            rw = range.Rows.Count;
            cl = range.Columns.Count;


            for (rCnt = 1; rCnt <= rw; rCnt++)
            {
                for (cCnt = 1; cCnt <= cl; cCnt++)
                {
                    var val = (range.Cells[rCnt, cCnt] as Range).Value2;

                    Console.WriteLine(val);
                }
            }

            workbook.Close(true, null, null);
            application.Quit();

            Marshal.ReleaseComObject(worksheet);
            Marshal.ReleaseComObject(workbook);
            Marshal.ReleaseComObject(application);
        }


        static void DisplayInExcel(IEnumerable<Box> boxes)
        {
            // Excel.Application
            var excelApp = new Application();
            excelApp.Visible = true;

            excelApp.Workbooks.Add();

            _Worksheet worksheet = (Worksheet)excelApp.ActiveSheet;

            worksheet.Cells[1, "A"] = "Kutu Genişliği";
            worksheet.Cells[1, "B"] = "Kutu Yüksekliği";
            worksheet.Cells[1, "C"] = "Kutu Eni";

            var row = 1;
            foreach (var box in boxes)
            {
                row++;

                worksheet.Cells[row, "A"] = box.width;
                worksheet.Cells[row, "B"] = box.height;
                worksheet.Cells[row, "C"] = box.tick;

                ((Range)worksheet.Columns[1]).AutoFit();
                ((Range)worksheet.Columns[2]).AutoFit();
                ((Range)worksheet.Columns[3]).AutoFit();
            }
        }


    }
}
