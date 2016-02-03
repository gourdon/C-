imports System.Net ' WebClient
imports System.Runtime.InteropServices ' Guid, ClassInterface, ClassInterfaceType
imports System.Globalization ' CultureInfo
 
<Guid("FC06752C-C091-402F-A902-80D3D58DC861")>
Public Interface IYahooAPIVBNet
    Function GetBid(symbol As String) As Double
    Function GetAsk(symbol As String) As Double
    Function GetCapitalization(symbol As String) As String
    Function GetValues(symbol As String, fields As String) As String()
End Interface

<Guid("4103B22F-5424-46D6-A960-3C1440D96316")>
<ClassInterface(ClassInterfaceType.None)>
Public Class YahooAPIVBNet
Implements IYahooAPIVBNet

    Private Shared ReadOnly webClient As WebClient = new WebClient()

    Private Const UrlTemplate As String = "http://finance.yahoo.com/d/quotes.csv?s={0}&f={1}"

    Private Shared Function ParseDouble(value As String) As Double
        return Double.Parse(value.Trim(), CultureInfo.InvariantCulture)
    End Function
    
    Private Shared Function GetDataFromYahoo(symbol As String, fields As String) As String()
        Dim request As String = String.Format(UrlTemplate, symbol, fields)

        Dim rawData As String = webClient.DownloadString(request).Trim

        return rawData.Split(New [Char]() {","})
    End Function
    
    Public Function GetBid(symbol As String) As Double Implements IYahooAPIVBNet.GetBid
        return ParseDouble(GetDataFromYahoo(symbol, "b3")(0))
    End Function

    Public Function GetAsk(symbol As String) As Double Implements IYahooAPIVBNet.GetAsk
        return ParseDouble(GetDataFromYahoo(Symbol, "b2")(0))
    End Function
    
    Public Function GetCapitalization(symbol As String) As String Implements IYahooAPIVBNet.GetCapitalization
        return GetDataFromYahoo(symbol, "j1")(0)
    End Function
    
    Public Function GetValues(symbol As String, fields As String) As String() Implements IYahooAPIVBNet.GetValues
        return GetDataFromYahoo(symbol, fields)
    End Function
End Class