using System.ComponentModel;
using System.Linq;
using Microsoft.Office.Interop.Excel;

namespace Chapter01VS2010
{
    [Description("Listing 1.19")]
    class SaveInExcel
    {
        static void Main()
        {
            var app = new Application { Visible = false };
            Workbook workbook = app.Workbooks.Add();
            Worksheet worksheet = app.ActiveSheet;
            int row = 1;
            foreach (var product in Product.GetSampleProducts()
                                           .Where(p => p.Price != null))
            {
                worksheet.Cells[row, 1].Value = product.Name;
                worksheet.Cells[row, 2].Value = product.Price;
                row++;
            }
            workbook.SaveAs(Filename: "demo.xls",
                            FileFormat: XlFileFormat.xlWorkbookNormal);
            app.Application.Quit();
        }
    }
}
