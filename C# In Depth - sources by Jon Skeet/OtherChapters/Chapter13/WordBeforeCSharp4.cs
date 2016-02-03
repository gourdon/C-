using System;
using Microsoft.Office.Interop.Word;
using System.ComponentModel;

namespace Chapter13
{
    [Description("Listing 13.08")]
    class WordBeforeCSharp4
    {
        static void Main()
        {
            object missing = Type.Missing;
            var app = new Application { Visible = false };
            app.Documents.Add(ref missing, ref missing, ref missing, ref missing);
            Document doc = app.ActiveDocument;
            Paragraph para = doc.Paragraphs.Add(ref missing);
            para.Range.Text = "Thank goodness for C# 4";
            object filename = "demo.doc";
            object format = WdSaveFormat.wdFormatDocument97;
            doc.SaveAs(ref filename, ref format,
                       ref missing, ref missing, ref missing,
                       ref missing, ref missing, ref missing,
                       ref missing, ref missing, ref missing,
                       ref missing, ref missing, ref missing,
                       ref missing, ref missing);
            doc.Close(ref missing, ref missing, ref missing);
            app.Application.Quit(ref missing, ref missing, ref missing);
        }
    }
}
