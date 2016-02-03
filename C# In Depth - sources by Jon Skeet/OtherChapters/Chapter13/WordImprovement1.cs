using Microsoft.Office.Interop.Word;
using System.ComponentModel;

namespace Chapter13
{
    [Description("Listing 13.09")]
    class WordImprovement1
    {
        static void Main()
        {
            var app = new Application { Visible = false };
            app.Documents.Add();
            Document doc = app.ActiveDocument;            
            Paragraph para = doc.Paragraphs.Add();
            para.Range.Text = "Thank goodness for C# 4";

            object filename = "demo.doc";
            object format = WdSaveFormat.wdFormatDocument97;
            doc.SaveAs(FileName: ref filename, FileFormat: ref format);
            doc.Close();
            app.Application.Quit();
        }
    }
}
