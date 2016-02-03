using System.Collections.Generic;

namespace Chapter01
{
    class ProductWithNullablePrice
    {
        public string Name { get; private set; }
        public decimal? Price { get; private set; }

        public ProductWithNullablePrice(string name, decimal price)
        {
            Name = name;
            Price = price;
        }

        ProductWithNullablePrice()
        {
        }

        public static List<ProductWithNullablePrice> GetSampleProducts()
        {
            return new List<ProductWithNullablePrice>
            {
                new ProductWithNullablePrice { Name="West Side Story", Price = 9.99m },
                new ProductWithNullablePrice { Name="Assassins", Price=14.99m },
                new ProductWithNullablePrice { Name="Frogs", Price=13.99m },
                new ProductWithNullablePrice { Name="Sweeney Todd", Price=null}
            };
        }

        public override string ToString()
        {
            return string.Format("{0}: {1}", Name, Price);
        }
    }
}
