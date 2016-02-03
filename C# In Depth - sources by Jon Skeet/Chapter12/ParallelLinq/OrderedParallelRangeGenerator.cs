using System.Linq;

namespace ParallelLinq
{
    class OrderedParallelRangeGenerator : MandelbrotGenerator
    {
        private OrderedParallelRangeGenerator(ImageOptions options)
            : base(options)
        {
        }

        static void Main()
        {
            var generator = new OrderedParallelRangeGenerator(ImageOptions.Default);
            generator.Display();
        }

        protected override byte[]  GeneratePixels()
        {
            var query = from row in ParallelEnumerable.Range(0, Height).AsOrdered()
                        from column in Enumerable.Range(0, Width)
                        select ComputeIndex(row, column);

            return query.ToArray();
        }
    }
}
