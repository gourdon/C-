using System;
using System.Collections;
using System.ComponentModel;

namespace Chapter01.CSharp1
{
    [Description("Listing 1.10")]
    class ArrayListQuery
    {
        static void Main()
        {
            ArrayList products = Product.GetSampleProducts();
            foreach (Product product in products)
            {
                if (product.Price > 10m)
                {
                    Console.WriteLine(product);
                }
            }
        }
    }
}
