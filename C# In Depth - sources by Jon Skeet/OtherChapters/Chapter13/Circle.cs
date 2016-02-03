using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows;

namespace Chapter13
{
    public sealed class Circle : IShape
    {
        private readonly Point center;
        public Point Center { get { return center; } }

        private readonly double radius;
        public double Radius { get { return radius; } }

        public Circle(Point center, int radius)
        {
            this.center = center;
            this.radius = radius;
        }

        public double Area
        {
            get { return Math.PI * radius * radius; }
        }

        public Rect BoundingBox
        {
            get
            {
                return new Rect(center - new Vector(radius, radius),
                                new Size(radius * 2, radius * 2));
            }
        }
    }
}
