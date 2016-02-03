using ContractAssemblyDemo;

namespace Chapter15
{
    class ContractAssemblyDemoTest
    {
        static void Main()
        {
            new PreconditionDemo().DontPassInNull(null);
            new PostconditionDemo().GimmeFive();
        }
    }
}
