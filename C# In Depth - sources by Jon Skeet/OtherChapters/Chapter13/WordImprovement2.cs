using System;
using System.Collections.Generic;
using Microsoft.Office.Interop.Word;
using System.Linq;
using System.Text;
using System.ComponentModel;

namespace Chapter13
{
    [Description("Listing 13.10")]
    class WordImprovement2
    {
        static void Main()
        {
            var app = new Application { Visible = false };
            app.Documents.Add();
            Document doc = app.ActiveDocument;
            Paragraph para = doc.Paragraphs.Add();
            para.Range.Text = "Thank goodness for C# 4";
            doc.SaveAs(FileName: "test.doc",
                       FileFormat: WdSaveFormat.wdFormatDocument97);
            doc.Close();
            app.Application.Quit();
        }
    }
}
