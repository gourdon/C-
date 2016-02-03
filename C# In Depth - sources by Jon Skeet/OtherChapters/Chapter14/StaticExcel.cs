using System.ComponentModel;
using System.Linq;
using Microsoft.Office.Interop.Excel;

namespace Chapter14
{
    [Description("Listing 14.05")]
    class StaticExcel
    {
        static void Main()
        {
            var app = new Application { Visible = true };
            app.Workbooks.Add();
            Worksheet worksheet = (Worksheet) app.ActiveSheet;
            Range start = (Range) worksheet.Cells[1, 1];
            Range end = (Range) worksheet.Cells[1, 20];
            worksheet.Range[start, end].Value = Enumerable.Range(1, 20).ToArray();
        }
    }
}
