using System.ComponentModel;
using System.Linq;
using Microsoft.Office.Interop.Excel;

namespace Chapter14
{
    [Description("Listing 14.07")]
    class VeryDynamicExcel
    {
        static void Main()
        {
            var app = new Application { Visible = true };
            app.Workbooks.Add();
            dynamic worksheet = app.ActiveSheet;
            dynamic start = worksheet.Cells[1, 1];
            dynamic end = worksheet.Cells[1, 20];
            worksheet.Range[start, end].Value = Enumerable.Range(1, 20)
                                                          .ToArray();
        }
    }
}
