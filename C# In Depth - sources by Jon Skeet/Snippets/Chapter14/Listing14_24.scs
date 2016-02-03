var list = new List<dynamic> { 50, 5m, 5d, 3 };
var query = from number in list
            where number > 4
            select (number / 20) * 10;

foreach (var item in query)
{
    Console.WriteLine(item);
}