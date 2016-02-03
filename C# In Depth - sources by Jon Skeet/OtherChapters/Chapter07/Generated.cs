using System;

namespace Chapter07
{
    partial class PartialMethodDemo
    {
        public PartialMethodDemo()
        {
            OnConstructorStart();
            Console.WriteLine("Generated constructor");
            OnConstructorEnd();
        }

        partial void OnConstructorStart();
        partial void OnConstructorEnd();
    }
}