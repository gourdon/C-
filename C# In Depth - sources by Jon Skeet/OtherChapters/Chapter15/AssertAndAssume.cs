using System;
using System.Diagnostics.Contracts;

namespace Chapter15
{
    class AssertAndAssume
    {
        public static int RollDice(Random rng)
        {
            Contract.Ensures(Contract.Result<int>() >= 2);
            Contract.Ensures(Contract.Result<int>() <= 12);

            if (rng == null)
            {
                rng = new Random();
            }
            Contract.Assert(rng != null);

            int firstRoll = rng.Next(1, 7);
            Contract.Assume(firstRoll >= 1);
            Contract.Assume(firstRoll <= 6);
            int secondRoll = rng.Next(1, 7);
            Contract.Assume(secondRoll >= 1);
            Contract.Assume(secondRoll <= 6);

            return firstRoll + secondRoll;
        }

        static void Main()
        {
            Console.WriteLine(RollDice(null));
            Console.WriteLine(RollDice(new Random(10)));
        }

    }
}
