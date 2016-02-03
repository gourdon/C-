using System.Diagnostics.Contracts;

namespace Chapter15
{
    class ComplexPostConditionFixed
    {
        static bool TryParsePreserveValue(string text, ref int value)
        {
            Contract.Ensures(Contract.Result<bool>() ||
                Contract.OldValue(value) == Contract.ValueAtReturn(out value));
            int temp;
            if (int.TryParse(text, out temp))
            {
                value = temp;
                return true;
            }
            return false;
        }

        static void Main()
        {
            int value = 10;
            TryParsePreserveValue("20", ref value);
            TryParsePreserveValue("bad", ref value);
        }
    }
}
