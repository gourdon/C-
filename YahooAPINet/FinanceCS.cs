using System.Net; // WebClient
using System.Runtime.InteropServices; // Guid, ClassInterface, ClassInterfaceType
using System.Globalization; // CultureInfo
 
[Guid("97E1D9DB-8478-4E56-9D6D-26D8EF13B100")]
public interface IYahooAPICS
{
    double GetBid(string symbol);
    double GetAsk(string symbol);
    string GetCapitalization(string symbol);
    string[] GetValues(string symbol, string fields);
}

[Guid("BBF87E31-77E2-46B6-8093-1689A144BFC6")]
[ClassInterface(ClassInterfaceType.None)]
public class YahooAPICS : IYahooAPICS
{
    private static readonly WebClient webClient = new WebClient();

    private const string UrlTemplate = "http://finance.yahoo.com/d/quotes.csv?s={0}&f={1}";

    private static double ParseDouble(string value)
    {
         return double.Parse(value.Trim(), CultureInfo.InvariantCulture);
    }
    
    private static string[] GetDataFromYahoo(string symbol, string fields)
    {
        string request = string.Format(UrlTemplate, symbol, fields);

        string rawData = webClient.DownloadString(request).Trim();
        
        return rawData.Split(',');
    }

    public double GetBid(string symbol)
    {
        return ParseDouble(GetDataFromYahoo(symbol, "b3")[0]);
    }

    public double GetAsk(string symbol)
    {
        return ParseDouble(GetDataFromYahoo(symbol, "b2")[0]);
    }
    
    public string GetCapitalization(string symbol)
    {
        return GetDataFromYahoo(symbol, "j1")[0];
    }
    
    public string[] GetValues(string symbol, string fields)
    {
        return GetDataFromYahoo(symbol, fields);
    }
}