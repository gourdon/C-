using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;

namespace Chapter01
{
    [Description("Listing 1.09")]
    class ListOrderWithExtensionMethod
    {
        static void Main()
        {
            List<Product> products = Product.GetSampleProducts();

            foreach (Product product in products.OrderBy(p => p.Name))
            {
                Console.WriteLine(product);
            }
        }
    }
}
