using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.ComponentModel;

namespace AppendixB
{
    [Description("Listing B.3")]
    class StackIteration
    {
        static void Main()
        {
            Stack<int> stack = new Stack<int>();
            stack.Push(1);
            stack.Push(2);
            stack.Push(3);
            foreach (int value in stack)
            {
                Console.WriteLine(value);
            }
        }
    }
}
