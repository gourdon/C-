using System;
using System.Collections.Generic;
using System.ComponentModel;
using MiscUtil.Collections;
using MiscUtil.Collections.Extensions;

namespace Chapter09
{
    [Description("Demonstration of ProjectionComparer (from MiscUtil)")]
    class ProductComparisonDemo2
    {
        static void Main()
        {
            // This uses lambda expressions and extension methods...
            IComparer<Product> comparer = ProjectionComparer<Product>
                .Create(p => p.Popularity).Reverse()
                .ThenBy(p => p.Price)
                .ThenBy(p => p.Name);

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
    }
}
