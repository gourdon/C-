using System;
using System.ComponentModel;

namespace Chapter09
{
    [Description("Listing 9.11")]
    class ConversionWithTypeInference
    {
        static void PrintConvertedValue<TInput, TOutput>
            (TInput input, Converter<TInput, TOutput> converter)
        {
            Console.WriteLine(converter(input));
        }

        static void Main()
        {
            PrintConvertedValue("I'm a string", x => x.Length);
        }
    }
}
