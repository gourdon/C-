using System.Collections.Generic;
using System.ComponentModel;

namespace Chapter01VS2010
{
    [Description("Listing 1.04 (and more)")]
    public class Product
    {
        readonly string name;
        public string Name { get { return name; } }

        decimal? price;
        public decimal? Price { get { return price; } }

        public Product(string name, decimal? price = null)
        {
            this.name = name;
            this.price = price;
        }

        public static List<Product> GetSampleProducts()
        {
            List<Product> list = new List<Product>();
            list.Add(new Product(name: "West Side Story", price: 9.99m));
            list.Add(new Product(name: "Assassins", price: 14.99m));
            list.Add(new Product(name: "Frogs", price: 13.99m));
            list.Add(new Product(name: "Sweeney Todd", price: 10.99m));
            list.Add(new Product(name: "Unpriced"));
            return list;
        }

        public override string ToString()
        {
            return string.Format("{0}: {1}", name, price);
        }
    }
}