using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;

namespace Chapter03.Variance
{
    [Description("Extra code for P105 (variance and comparisons)")]
    class ShapeSorting
    {
        static void Main()
        {
            IComparer<IShape> areaComparer = new AreaComparer();
            List<Circle> circles = new List<Circle>();
            circles.Add(new Circle(Point.Empty, 20));
            circles.Add(new Circle(Point.Empty, 10));
            circles.Add(new Circle(Point.Empty, 40));
            circles.Add(new Circle(Point.Empty, 30));

            // Won't work - wrong kind of comparer
            // circles.Sort(areaComparer);

            // Adapt the "general" comparer to work with more specific values
            IComparer<Circle> circleComparer = new ComparisonHelper<IShape, Circle>(areaComparer);
            circles.Sort(circleComparer);

            foreach (Circle circle in circles)
            {
                Console.WriteLine(circle.Area);
            }
        }
    }
}
