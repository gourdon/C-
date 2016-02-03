using System;
using System.Collections.Generic;
using System.ComponentModel;

namespace Chapter01.CSharp2
{
    [Description("Listing 1.12")]
    class ListQueryWithDelegatesCompact
    {
        static void Main()
        {
            List<Product> products = Product.GetSampleProducts();
            products.FindAll(delegate(Product p) { return p.Price > 10; })
                    .ForEach(delegate(Product p) { Console.WriteLine(p); });
        }
    }
}
