using System.ComponentModel;
using IronPython.Hosting;
using Microsoft.Scripting.Hosting;

namespace Chapter14
{
    [Description("Listing 14.08")]
    class EmbeddedHelloWorld
    {
        static void Main()
        {
            ScriptEngine engine = Python.CreateEngine();
            engine.Execute("print 'hello, world'");
            engine.ExecuteFile("HelloWorld.py");
        }
    }
}
