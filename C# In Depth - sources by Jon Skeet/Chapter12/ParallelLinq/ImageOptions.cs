using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Imaging;
using System.Linq;
using System.Text;

namespace ParallelLinq
{
    public class ImageOptions
    {
        private readonly int maxIterations;
        public int MaxIterations { get { return maxIterations; } }

        private readonly int imageWidth;
        public int ImageWidth { get { return imageWidth; } }
            
        private readonly int imageHeight;
        public int ImageHeight { get { return imageHeight; } }

        private readonly double sampleWidth = 3.2;
        public double SampleWidth { get { return sampleWidth; } }

        private readonly double sampleHeight = 2.5;
        public double SampleHeight { get { return sampleHeight; } }

        private readonly double offsetX = -2.1;
        public double OffsetX { get { return offsetX; } }

        private readonly double offsetY = -1.25;
        public double OffsetY { get { return offsetY; } }

        public static readonly ImageOptions Default = new ImageOptions(4000, 1024, 3.2, 2.5, -2.1, -1.25);

        public ImageOptions(int maxIterations,
            int imageWidth,
            double sampleWidth,
            double sampleHeight,
            double offsetX,
            double offsetY)
        {
            this.maxIterations = maxIterations;
            this.imageWidth = imageWidth;
            this.sampleWidth = sampleWidth;
            this.sampleHeight = sampleHeight;
            this.offsetX = offsetX;
            this.offsetY = offsetY;
            this.imageHeight = (int) (sampleHeight * imageWidth / sampleWidth);
        }
    }
}
