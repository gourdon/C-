using System;
using System.ComponentModel;
using System.Reflection;

namespace Chapter03
{
    [Description("Listing 3.13")]
    class GenericMethodReflection
    {
        public static void PrintTypeParameter<T>()
        {
            Console.WriteLine (typeof(T));
        }

        static void Main()
        {
            Type type = typeof(GenericMethodReflection);
            MethodInfo definition = type.GetMethod("PrintTypeParameter");        
            MethodInfo constructed;
            constructed = definition.MakeGenericMethod(typeof(string));       
            constructed.Invoke(null, null);
        }
    }
}
