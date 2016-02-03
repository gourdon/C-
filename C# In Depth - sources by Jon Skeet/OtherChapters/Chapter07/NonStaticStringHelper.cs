using System;
using System.ComponentModel;

namespace Chapter07
{
    [Description("Listing 7.03")]
    public sealed class NonStaticStringHelper             
    {
        private NonStaticStringHelper()                    
        {
        }
        
        public static string Reverse(string input)
        {
            char[] chars = input.ToCharArray();
            Array.Reverse(chars);
            return new string(chars);
        }
    }
}
