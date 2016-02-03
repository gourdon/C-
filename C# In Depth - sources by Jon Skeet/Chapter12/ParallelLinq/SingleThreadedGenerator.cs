using System.Linq;
using System.ComponentModel;

namespace ParallelLinq
{
    [Description("Listing 12.10")]
    class SingleThreadedGenerator : MandelbrotGenerator
    {
        private SingleThreadedGenerator(ImageOptions options)
            : base(options)
        {
        }

        static void Main()
        {
            var generator = new SingleThreadedGenerator(ImageOptions.Default);
            generator.Display();
        }

        protected override byte[]  GeneratePixels()
        {
            var query = from row in Enumerable.Range(0, Height)
                        from column in Enumerable.Range(0, Width)
                        select ComputeIndex(row, column);

            return query.ToArray();
        }
    }
}
