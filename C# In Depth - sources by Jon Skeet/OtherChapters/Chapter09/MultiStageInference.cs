using System;
using System.ComponentModel;

namespace Chapter09
{
    [Description("Listing 9.15")]
    class MultiStageInference
    {
        static void ConvertTwice<TInput, TMiddle, TOutput> (TInput input,
                                                            Converter<TInput, TMiddle> firstConversion,
                                                            Converter<TMiddle, TOutput> secondConversion)
        {
            TMiddle middle = firstConversion(input);
            TOutput output = secondConversion(middle);
            Console.WriteLine(output);
        }

        static void Main()
        {
            ConvertTwice("Another string",
                         text => text.Length,
                         length => Math.Sqrt(length));
        }
    }
}
