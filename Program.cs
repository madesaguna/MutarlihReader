using System;

namespace ImportPemutakhiran
{
    class Program
    {

        static void Main(string[] args)
        {
            var x = new ExcelReader(@"dpshp.xlsx");
            x.Execute();
            //Set the cell value using row and column.


            //The style object is used to access most cells formatting and styles.
            //ws.Cells[3, 1].Style.Font.Bold = true;
            //p.Save();

            Console.WriteLine("--DONE--");
        }
    }
}
