using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;

namespace Chapter04
{
    [Description("Demonstration of PartialComparer")]
    class ProductComparisonDemo : IComparer<Product>, IEqualityComparer<Product>
    {
        static void Main()
        {
            // See ProductComparisonDemo2 in Chapter09 for an even nicer solution :)
            IComparer<Product> comparer = new ProductComparisonDemo();

            // Products are sorted by popularity (highest first), price (lowest first) and then name
            Product remote = new Product(100, 25.00m, "Remote control");
            Product mouse = new Product(50, 10.00m, "Wireless mouse");
            Product headphones = new Product(100, 20.00m, "USB drive");
            Product wipes = new Product(50, 10.00m, "Screen wipes");
            Product game = new Product(150, 40.00m, "Shiny new game");

            Product[] products = { remote, mouse, headphones, wipes, game };

            Array.Sort(products, comparer);
            foreach (Product product in products)
            {
                Console.WriteLine(product);
            }

            Product otherGame = new Product(150, 40.00m, "Shiny new game");
            Console.WriteLine("Two similar products compare as 0? {0}", comparer.Compare(game, otherGame) == 0);
        }

        public int Compare(Product first, Product second)
        {
            return PartialComparer.ReferenceCompare(first, second) ??
                   // Reverse comparison of popularity to sort descending
                   PartialComparer.Compare(second.Popularity, first.Popularity) ??
                   PartialComparer.Compare(first.Price, second.Price) ??
                   PartialComparer.Compare(first.Name, second.Name) ??
                   0;
        }

        public bool Equals(Product x, Product y)
        {
            // This gives an "early out" for equal references or the case where one reference is null.
            return PartialComparer.ReferenceEquals(x, y) ??
                x.Name == y.Name &&
                x.Popularity == y.Popularity &&
                x.Price == y.Price;
        }

        public int GetHashCode(Product obj)
        {
            throw new NotImplementedException();
        }
    }
}
