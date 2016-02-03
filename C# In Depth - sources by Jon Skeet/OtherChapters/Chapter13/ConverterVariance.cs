using System;
using System.ComponentModel;

namespace Chapter13
{
    [Description("Listing 13.15")]
    class ConverterVariance
    {
        static void Main()
        {
            Converter<object, string> converter = x => x.ToString();
            Converter<string, string> contravariance = converter;
            Converter<object, object> covariance = converter;
            Converter<string, object> both = converter; 
        }        
    }
}
