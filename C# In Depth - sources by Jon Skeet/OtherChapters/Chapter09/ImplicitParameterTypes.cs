using System;
using System.ComponentModel;

namespace Chapter09
{
    [Description("Implicitly typed parameter lists")]
    class ImplicitParameterTypes
    {
        // None of the built-in delegates have out parameters, so we have to declare our own
        delegate void OutAction(out int x);

        static void Main()
        {
            // Just declare the variables first to make the examples consistent
            Func<string, string, int> func;
            OutAction outAction;

            // Valid: both parameters implicit
            func = (x, y) => x.Length + y.Length;

            // Valid: both parameters explicit
            func = (string x, string y) => x.Length + y.Length;

            // Invalid: trying to make one parameter implicit and one explicit
            // func = (string x, y) => x.Length + y.Length;

            // Valid: explicitly typed out parameter
            outAction = (out int x) => { x = 10; };

            // Invalid: out (and ref) parameters can't be implicit. You can't
            // even just tell the compiler that it's an out parameter and let it infer the type.
            // outAction = (out x) => { x = 10; };
            // outAction = (x) => { x = 10; };
        }
    }
}
