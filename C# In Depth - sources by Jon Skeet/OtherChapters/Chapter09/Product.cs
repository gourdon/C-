namespace Chapter09
{
    /// <summary>
    /// Simple data class to compound comparisons
    /// </summary>
    public sealed class Product
    {
        private readonly int popularity;
        private readonly decimal price;
        private readonly string name;

        public int Popularity { get { return popularity; } }
        public decimal Price { get { return price; } }
        public string Name { get { return name; } }

        public Product(int popularity, decimal price, string name)
        {
            this.popularity = popularity;
            this.price = price;
            this.name = name;
        }

        public override string ToString()
        {
            return string.Format("{0}: Price={1:c}, popularity={2}", Name, Price, Popularity);
        }
    }
}
