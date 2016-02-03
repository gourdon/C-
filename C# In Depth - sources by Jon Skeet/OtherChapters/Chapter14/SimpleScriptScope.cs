using System;
using System.ComponentModel;
using IronPython.Hosting;
using Microsoft.Scripting.Hosting;

namespace Chapter14
{
    [Description("Listing 14.09")]
    class SimpleScriptScope
    {
        static void Main()
        {
            string python = @"
text = 'hello'
output = input + 1
";
            ScriptEngine engine = Python.CreateEngine();
            ScriptScope scope = engine.CreateScope();
            scope.SetVariable("input", 10);
            engine.Execute(python, scope);
            Console.WriteLine(scope.GetVariable("text"));
            Console.WriteLine(scope.GetVariable("input"));
            Console.WriteLine(scope.GetVariable("output"));
        }
    }
}
