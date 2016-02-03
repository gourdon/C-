using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;

namespace Chapter01
{
    [Description("Listing 1.14")]
    class DisplayProductsWithUnknownPrice
    {
        static void Main()
        {
            List<ProductWithNullablePrice> products = ProductWithNullablePrice.GetSampleProducts();
            foreach (ProductWithNullablePrice product in products.Where(p => p.Price == null))
            {
                Console.WriteLine(product.Name);
            }
        }
    }
}
