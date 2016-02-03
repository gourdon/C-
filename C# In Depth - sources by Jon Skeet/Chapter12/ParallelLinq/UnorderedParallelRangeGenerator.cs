using System.Linq;

namespace ParallelLinq
{
    class UnorderedParallelRangeGenerator : MandelbrotGenerator
    {
        private UnorderedParallelRangeGenerator(ImageOptions options)
            : base(options)
        {
        }

        static void Main()
        {
            var generator = new UnorderedParallelRangeGenerator(ImageOptions.Default);
            generator.Display();
        }

        protected override byte[]  GeneratePixels()
        {
            var query = from row in ParallelEnumerable.Range(0, Height)
                        from column in Enumerable.Range(0, Width)
                        select ComputeIndex(row, column);

            return query.ToArray();
        }
    }
}
