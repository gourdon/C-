using System;
using System.ComponentModel;

namespace Chapter05
{
    [Description("Listing 5.04")]
    class BreakingChange
    {
        delegate void SampleDelegate(string x);

        public void CandidateAction(string x)
        {
            Console.WriteLine("Snippet.CandidateAction");
        }

        public class Derived : BreakingChange
        {
            public void CandidateAction(object o)
            {
                Console.WriteLine("Derived.CandidateAction");
            }
        }

        static void Main()
        {
            Derived x = new Derived();
            SampleDelegate factory = new SampleDelegate(x.CandidateAction);
            factory("test");
        }
    }
}
