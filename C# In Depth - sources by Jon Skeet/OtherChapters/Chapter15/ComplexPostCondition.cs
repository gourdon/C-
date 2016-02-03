using System.Diagnostics.Contracts;

namespace Chapter15
{
    class ComplexPostCondition
    {
        static bool TryParsePreserveValue(string text, ref int value)
        {
            Contract.Ensures(Contract.Result<bool>() ||
                Contract.OldValue(value) == Contract.ValueAtReturn(out value));
            return int.TryParse(text, out value);
        }

        static void Main()
        {
            int value = 10;
            TryParsePreserveValue("20", ref value);
            TryParsePreserveValue("bad", ref value);
        }
    }
}
