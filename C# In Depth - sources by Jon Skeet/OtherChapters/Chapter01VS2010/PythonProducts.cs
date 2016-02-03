using IronPython.Hosting;
using Microsoft.Scripting.Hosting;
using Microsoft.Office.Interop.Excel;
using System.ComponentModel;
using System;

namespace Chapter01VS2010
{
    [Description("Listing 1.20")]
    class PythonProducts
    {
        static void Main()
        {
            ScriptEngine engine = Python.CreateEngine();
            ScriptScope scope = engine.ExecuteFile("FindProducts.py");
            dynamic products = scope.GetVariable("products");
            foreach (dynamic product in products)
            {
                Console.WriteLine("{0}: {1}", product.ProductName, product.Price);
            }
        }
    }
}
