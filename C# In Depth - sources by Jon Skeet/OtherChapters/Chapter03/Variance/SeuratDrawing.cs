using System;
using System.Collections.Generic;

namespace Chapter03.Variance
{
    public sealed class SeuratDrawing : IDrawing
    {
        public IEnumerable<IShape> Shapes
        {
            get { throw new NotImplementedException(); }
        }
    }
}
