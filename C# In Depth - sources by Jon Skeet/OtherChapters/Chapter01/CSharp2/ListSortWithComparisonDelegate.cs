using System;
using System.Collections.Generic;
using System.ComponentModel;

namespace Chapter01.CSharp2
{
    [Description("Listing 1.07")]
    class ListSortWithComparisonDelegate
    {
        static void Main()
        {
            List<Product> products = Product.GetSampleProducts();
            products.Sort(delegate(Product first, Product second)
                { return first.Name.CompareTo(second.Name); }
            );
            foreach (Product product in products)
            {
                Console.WriteLine(product);
            }
        }
    }
}
