using System.ComponentModel;
using System.Linq;

namespace Chapter14
{
    [Description("Listing 14.22")]
    class DynamicExtensionMethods
    {
        static void Main()
        {
            dynamic size = 5;
            var numbers = Enumerable.Range(10, 10);
            // Error: extension methods can't be dynamically dispatched
            // var error = numbers.Take(size);
            var workaround1 = numbers.Take((int) size);
            var workaround2 = Enumerable.Take(numbers, size);
        }
    }
}
