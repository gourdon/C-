
namespace Chapter15
{
    class MoreArrayBoundsChecking
    {
        static void Main(string[] args)
        {
            string[] copy = new string[args.Length * 2 - 1];
            for (int i = 0; i < args.Length; i++)
            {
                copy[i * 2] = args[i];
            }
        }
    }
}
