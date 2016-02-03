using System;
using System.Drawing;

namespace Chapter03.Variance
{
    public sealed class Circle : IShape
    {
        private readonly double radius;
        private readonly Point center;

        public Circle(Point center, double radius)
        {
            this.radius = radius;
            this.center = center;
        }

        public double Area
        {
            get { return Math.PI * radius * radius; }
        }
    }
}
