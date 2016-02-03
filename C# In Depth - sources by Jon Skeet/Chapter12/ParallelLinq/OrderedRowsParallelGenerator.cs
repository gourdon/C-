using System.ComponentModel;
using System.Linq;

namespace ParallelLinq
{
    [Description("Listing 12.12")]
    class OrderedRowsParallelGenerator : MandelbrotGenerator
    {
        private OrderedRowsParallelGenerator(ImageOptions options)
            : base(options)
        {
        }

        static void Main()
        {
            var generator = new OrderedRowsParallelGenerator(ImageOptions.Default);
            generator.Display();
        }

        protected override byte[]  GeneratePixels()
        {
            var query = from row in Enumerable.Range(0, Height)
                                              .AsParallel()
                                              .AsOrdered()
                        from column in Enumerable.Range(0, Width)
                        select ComputeIndex(row, column);

            return query.ToArray();
        }
    }
}
