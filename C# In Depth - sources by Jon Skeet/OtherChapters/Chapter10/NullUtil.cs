using System.ComponentModel;

namespace Chapter10
{
    [Description("Listing 10.05")]
    static class NullUtil
    {
        public static bool IsNull(this object x)
        {
            return x == null;
        }
    }
}
