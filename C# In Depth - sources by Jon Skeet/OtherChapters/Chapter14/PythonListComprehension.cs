using System;
using System.Collections.Generic;
using IronPython.Hosting;
using Microsoft.Scripting.Hosting;

namespace Chapter14
{
    class PythonListComprehension
    {
        static IEnumerable<int> ZeroToTen()
        {
            for (int i = 0; i <= 10; i++)
            {
                Console.WriteLine("Yielding {0}", i);
                yield return i;
            }
        }

        static void Main()
        {
            string python = @"
def executeOnEven(list, action):
# Note: replace () with [] for eager evaluation
    evenList = (item for item in list if item % 2 == 0)
    for item in evenList:
        action(item)
";
            ScriptEngine engine = Python.CreateEngine();
            ScriptScope scope = engine.CreateScope();
            engine.Execute(python, scope);
            dynamic function = scope.GetVariable("executeOnEven");
            IEnumerable<int> numbers = ZeroToTen();
            Action<int> print = x => Console.WriteLine("Received {0}", x);
            function(numbers, print);
        }
    }
}
