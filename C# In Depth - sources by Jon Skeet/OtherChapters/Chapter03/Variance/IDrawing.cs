using System.Collections.Generic;

namespace Chapter03.Variance
{
    public interface IDrawing
    {
        IEnumerable<IShape> Shapes { get; }
    }
}
