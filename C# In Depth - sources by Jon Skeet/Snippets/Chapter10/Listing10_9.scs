var collection = Enumerable.Range(0, 10)
                           .Where(x => x % 2 != 0)
                           .Reverse()
                           .Select(x => new { Original = x, SquareRoot = Math.Sqrt(x) });

foreach (var element in collection)
{
    Console.WriteLine("sqrt({0})={1}",
                      element.Original,
                      element.SquareRoot);
}