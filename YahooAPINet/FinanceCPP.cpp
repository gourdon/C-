#using <system.dll>
using namespace System; // String
using namespace System::Net; // WebClient
using namespace System::Runtime::InteropServices; // Guid, ClassInterface, ClassInterfaceType
using namespace System::Globalization; // CultureInfo
 
[Guid("4ABC0C71-9CF5-4618-8D7C-55E32DCF6314")]
public interface class IYahooAPICPP
{
    double GetBid(String^ symbol);
    double GetAsk(String^ symbol);
    String^ GetCapitalization(String^ symbol);
    cli::array<String^>^ GetValues(String^ symbol, String^ fields);
};

[Guid("AEC520AE-12D8-49A9-A5F4-853112C3B6AD")]
[ClassInterface(ClassInterfaceType::None)]
public ref class YahooAPICPP : IYahooAPICPP
{
    private: static initonly WebClient^ webClient = gcnew WebClient();

    private: literal String^ UrlTemplate = "http://finance.yahoo.com/d/quotes.csv?s={0}&f={1}";

	private: static double ParseDouble(String^ s)
	{
		return double::Parse(s->Trim(), CultureInfo::InvariantCulture);
	}
	
    private: static cli::array<String^>^ GetDataFromYahoo(String^ symbol, String^ fields)
    {
        String^ request = String::Format(UrlTemplate, symbol, fields);

        String^ rawData = webClient->DownloadString(request)->Trim();

        return  rawData->Split(',');
    }

    public: virtual double GetBid(String^ symbol)
    {
        return ParseDouble(GetDataFromYahoo(symbol, "b3")[0]);
    }

    public: virtual double GetAsk(String^ symbol)
    {
        return ParseDouble(GetDataFromYahoo(symbol, "b2")[0]);
    }
    
    public: virtual String^ GetCapitalization(String^ symbol)
    {
        return GetDataFromYahoo(symbol, "j1")[0];
    }
    
    public: virtual cli::array<String^>^ GetValues(String^ symbol, String^ fields)
    {
        return GetDataFromYahoo(symbol, fields);
    }
};