using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;

namespace Chapter01
{
    [Description("Listing 1.13")]
    class ListQueryWithLambdaExpression
    {
        static void Main()
        {
            List<Product> products = Product.GetSampleProducts();
            foreach (Product product in products.Where(p => p.Price > 10))
            {
                Console.WriteLine(product);
            }
        }
    }
}
