using System.ComponentModel;

namespace Chapter07
{
    [Description("Listing 7.10")]
    class UnusedWarningWithWarningSuppressed
    {
#pragma warning disable 0169
        int x;
#pragma warning restore 0169
    }
}
