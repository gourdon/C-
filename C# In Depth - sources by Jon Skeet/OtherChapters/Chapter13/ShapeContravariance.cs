using System.Collections.Generic;
using System.ComponentModel;

namespace Chapter13
{
    [Description("Listing 13.13")]
    class ShapeContravariance
    {
        class AreaComparer : IComparer<IShape>
        {
            public int Compare(IShape x, IShape y)
            {
                return x.Area.CompareTo(y.Area);
            }
        }

        static void Main()
        {
            IComparer<IShape> areaComparer = new AreaComparer();
            Shapes.Circles.Sort(areaComparer);
        }
    }
}
