using System;
using System.Collections.Generic;
using System.ComponentModel;

namespace Chapter01
{
    [Description("Listing 1.08")]
    class ListSortWithLambdaExpression
    {
        static void Main()
        {
            List<Product> products = Product.GetSampleProducts();
            products.Sort(
                (first, second) => first.Name.CompareTo(second.Name)
            );
            foreach (Product product in products)
            {
                Console.WriteLine(product);
            }
        }
    }
}
