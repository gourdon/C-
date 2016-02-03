using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Text.RegularExpressions;

namespace Chapter03
{
    [Description("Listing 3.01")]
    class DictionaryDemo
    {
        static Dictionary<string,int> CountWords(string text)
        {
            Dictionary<string,int> frequencies;
            frequencies = new Dictionary<string,int>();
                
            string[] words = Regex.Split(text, @"\W+");
                
            foreach (string word in words)
            {
                if (frequencies.ContainsKey(word))
                {
                    frequencies[word]++;
                }
                else
                {
                    frequencies[word] = 1;
                }
            }
            return frequencies;
        }

        static void Main()
        {
            string text = @"Do you like green eggs and ham?
                            I do not like them, Sam-I-am.
                            I do not like green eggs and ham.";

            Dictionary<string, int> frequencies = CountWords(text);
            foreach (KeyValuePair<string, int> entry in frequencies)
            {
                string word = entry.Key;
                int frequency = entry.Value;
                Console.WriteLine("{0}: {1}", word, frequency);
            }
        }
    }
}
