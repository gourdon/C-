using System;
using System.Collections.Generic;
using System.ComponentModel;

namespace Chapter01.CSharp2
{
    [Description("Listing 1.11")]
    class ListQueryWithDelegates
    {
        static void Main()
        {
            List<Product> products = Product.GetSampleProducts();
            Predicate<Product> test = delegate(Product p)
                { return p.Price > 10m; };
            List<Product> matches = products.FindAll(test);

            Action<Product> print = Console.WriteLine;
            matches.ForEach(print);
        }
    }
}
