using System;
using System.Collections.Generic;

namespace Chapter03.Variance
{
    public class MondrianDrawing : IDrawing
    {
        private readonly List<Rectangle> rectangles = new List<Rectangle>();

        // Other members as appropriate

        public IEnumerable<IShape> Shapes
        {
            get { throw new NotImplementedException(); }
        }
    }
}
