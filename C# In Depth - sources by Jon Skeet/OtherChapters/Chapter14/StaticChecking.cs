using System;
using System.ComponentModel;

namespace Chapter14
{
    [Description("Listing 14.20")]
    class StaticChecking
    {
        static void Main()
        {
            string text = "cut me up";
            dynamic guid = Guid.NewGuid();
            text.Substring(guid);
            // Compile-time error: no overload with string and something else
            // text.Substring("x", guid);
            // Compile-time error: no overload with three parameters
            // text.Substring(guid, guid, guid);
        }
    }
}
