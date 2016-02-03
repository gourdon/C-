using System;
using System.ComponentModel;
using Microsoft.Office.Interop.Word;

namespace Chapter13
{
    [Description("Listing 13.11")]
    class WordMeanings
    {
        static void ShowInfo(SynonymInfo info)
        {
            Console.WriteLine("{0} has {1} meanings",
               info.Word, info.MeaningCount);
        }

        static void Main()
        {
            var app = new Application { Visible = false };

            object missing = Type.Missing;
            ShowInfo(app.get_SynonymInfo("painful", ref missing));

            ShowInfo(app.SynonymInfo["nice", WdLanguageID.wdEnglishUS]);

            ShowInfo(app.SynonymInfo[Word: "features"]);

            app.Application.Quit();
        }
    }
}
