using System.ComponentModel;
namespace Chapter02
{
    [Description("Listing 2.3")]
    class ArrayVariance
    {
        static void Main()
        {
            string[] strings = new string[5];
            object[] objects = strings;
            objects[0] = new object();
        }
    }
}
