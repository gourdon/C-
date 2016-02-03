using System;
using System.ComponentModel;
using System.Windows;

namespace Chapter13
{
    [Description("Listing 13.14")]
    class SimpleDelegateVariance
    {
        static void Main()
        {
            Func<Square> squareFactory = () => new Square(new Point(5, 5), 10);
            Func<IShape> shapeFactory = squareFactory;

            Action<IShape> shapePrinter = shape => Console.WriteLine(shape.Area);
            Action<Square> squarePrinter = shapePrinter;

            squarePrinter(squareFactory());
            shapePrinter(shapeFactory());
        }
    }
}
