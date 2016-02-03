using System;
using System.ComponentModel;

namespace Chapter14
{
    [Description("Listing 14.21")]
    class TypeInferenceChecking
    {
        void Execute<T>(T first, T second, string other) where T : struct
        {
        }

        static void Main()
        {
            dynamic guid = Guid.NewGuid();
            Execute(10, 0, guid);
            // Error: No consistent value of T whatever the type of guid is
            // Execute(10, false, guid);
            // Error: T = string will be invalid whatever the type of guid is
            // Execute("hello", "hello", guid);
        }
    }
}
