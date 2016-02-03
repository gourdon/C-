using System.ComponentModel;
using IronPython.Hosting;
using Microsoft.Scripting.Hosting;

namespace Chapter14
{
    [Description("Listing 14.10")]
    class ScriptScopeDelegate
    {
        static void Main()
        {
            string python = @"
def sayHello(user):
    print 'Hello %(name)s' % {'name' : user}
";
            ScriptEngine engine = Python.CreateEngine();
            ScriptScope scope = engine.CreateScope();
            engine.Execute(python, scope);
            dynamic function = scope.GetVariable("sayHello");
            function("Jon");
        }
    }
}
