using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;

namespace Chapter01
{
    [Description("Listing 1.15")]
    class ListFilteringWithLinq
    {
        static void Main()
        {
            List<Product> products = Product.GetSampleProducts();
            var filtered = from Product p in products
                           where p.Price > 10
                           select p;
            foreach (Product product in filtered)
            {
                Console.WriteLine(product);
            }
        }
    }
}
