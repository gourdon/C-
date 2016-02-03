using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ParallelLinq
{
    class UnorderedPixelsParallelGenerator : MandelbrotGenerator
    {
        private UnorderedPixelsParallelGenerator(ImageOptions options)
            : base(options)
        {
        }

        static void Main()
        {
            var generator = new UnorderedPixelsParallelGenerator(ImageOptions.Default);
            generator.Display();
        }

        protected override byte[]  GeneratePixels()
        {
            var pixels = from row in Enumerable.Range(0, Height)
                         from column in Enumerable.Range(0, Width)
                         select new { row, column };
            var query = from pixel in pixels.AsParallel()
                        select ComputeIndex(pixel.row, pixel.column);

            return query.ToArray();
        }
    }
}
