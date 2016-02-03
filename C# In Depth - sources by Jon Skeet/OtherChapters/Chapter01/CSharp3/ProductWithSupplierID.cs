using System.Collections.Generic;

namespace Chapter01
{
    class ProductWithSupplierID
    {
        public string Name { get; private set; }
        public decimal Price { get; private set; }
        public int SupplierID { get; private set; }

        public ProductWithSupplierID(string name, decimal price)
        {
            Name = name;
            Price = price;
        }

        ProductWithSupplierID()
        {
        }

        public static List<ProductWithSupplierID> GetSampleProducts()
        {
            return new List<ProductWithSupplierID>
            {
                new ProductWithSupplierID { Name="West Side Story", Price = 9.99m, SupplierID=1 },
                new ProductWithSupplierID { Name="Assassins", Price=14.99m, SupplierID=2 },
                new ProductWithSupplierID { Name="Frogs", Price=13.99m, SupplierID=1 },
                new ProductWithSupplierID { Name="Sweeney Todd", Price=10.99m, SupplierID=3}
            };
        }

        public override string ToString()
        {
            return string.Format("{0}: {1}", Name, Price);
        }
    }
}
