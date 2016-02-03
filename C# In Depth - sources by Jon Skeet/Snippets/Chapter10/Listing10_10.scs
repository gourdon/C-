var collection = Enumerable.Range(-5, 11)
                           .Select(x => new { Original = x, Square = x * x })
                           .OrderBy(x => x.Square)
                           .ThenBy(x => x.Original);

foreach (var element in collection)
{
    Console.WriteLine(element);
}