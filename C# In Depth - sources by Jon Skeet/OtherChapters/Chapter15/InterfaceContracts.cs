using System.Diagnostics.Contracts;
using System.Globalization;

namespace Chapter15
{
    class InterfaceContracts
    {
        [ContractClass(typeof(ICaseConverterContracts))]
        public interface ICaseConverter
        {
            string Convert(string text);
        }

        [ContractClassFor(typeof(ICaseConverter))]
        private abstract class ICaseConverterContracts : ICaseConverter
        {
            public string Convert(string text)
            {
                Contract.Requires(text != null);
                Contract.Ensures(Contract.Result<string>() != null);
                return default(string);
            }

            private ICaseConverterContracts() {}
        }

        public class CurrentCultureUpperCaseFormatter : ICaseConverter
        {
            public string Convert(string text)
            {
                return text.ToUpper(CultureInfo.CurrentCulture);
            }
        }

        static void Main()
        {
            ICaseConverter converter = new CurrentCultureUpperCaseFormatter();
            converter.Convert(null);
        }
    }
}
