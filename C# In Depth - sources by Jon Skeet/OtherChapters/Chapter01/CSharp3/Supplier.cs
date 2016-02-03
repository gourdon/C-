using System.Collections.Generic;

namespace Chapter01
{
    class Supplier
    {
        public string Name { get; private set; }
        public int SupplierID { get; private set; }

        Supplier()
        {
        }

        public static List<Supplier> GetSampleSuppliers()
        {
            return new List<Supplier>
            {
                new Supplier { Name="Solely Sondheim", SupplierID=1 },
                new Supplier { Name="CD-by-CD-by-Sondheim", SupplierID=2 },
                new Supplier { Name="Barbershop CDs", SupplierID=3 }
            };
        }
    }
}