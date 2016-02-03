using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;

namespace Chapter13
{
    [Description("Listing 13.12")]
    class ShapeCovariance
    {
        static void Main()
        {
            List<IShape> shapesByAdding = new List<IShape>();
            shapesByAdding.AddRange(Shapes.Circles);
            shapesByAdding.AddRange(Shapes.Squares);

            List<IShape> shapesByConcat = Shapes.Circles.Concat<IShape>(Shapes.Squares).ToList();
        }
    }
}
