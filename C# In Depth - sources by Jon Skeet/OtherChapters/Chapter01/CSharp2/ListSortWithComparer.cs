using System;
using System.Collections.Generic;
using System.ComponentModel;

namespace Chapter01.CSharp2
{
    [Description("Listing 1.06")]
    class ListSortWithComparer
    {
        class ProductNameComparer : IComparer<Product>
        {
            public int Compare(Product first, Product second)
            {
                return first.Name.CompareTo(second.Name);
            }
        }

        static void Main()
        {
            List<Product> products = Product.GetSampleProducts();
            products.Sort(new ProductNameComparer());
            foreach (Product product in products)
            {
                Console.WriteLine(product);
            }
        }
    }
}
