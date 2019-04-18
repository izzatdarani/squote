Attribute VB_Name = "Squote"
' Squote module
Option Private Module ' disable Excel interface
Public Function GetStockQuote(ByVal stockTicker As String, ByVal startDate As String, ByVal endDate As String, ByVal Frequency As String) As String
    ' Get response function to obtain historical stock quotes from Yahoo! Finance
    
    ' init
    Dim cookie As String
    Dim crumb As String
    Dim objRequest_cc As Object
    Set objRequest_cc = CreateObject("WinHttp.WinHttpRequest.5.1")
    Dim api As String
    Dim stockQuotes As String
    Dim objRequest_yf As Object
    Set objRequest_yf = CreateObject("WinHttp.WinHttpRequest.5.1")

    ' get YF cookie crumb
    With objRequest_cc
        .Open "GET", "https://finance.yahoo.com/lookup?s=bananas", False
        .setRequestHeader "Content-Type", "application/x-www-form-urlencoded; charset=UTF-8"
        .Send
        ' get cookie
        cookie = Split(.getResponseHeader("Set-Cookie"), ";")(0)
        ' temporary crumb extraction
        ' please replace with more elegant parser in the future
        ' get crumb
        crumb = Mid(.ResponseText, InStrRev(.ResponseText, "crumb") + 8, 11)
        ' crumb is always 11 (not confirmed, only assumed based on pattern)
        If Len(crumb) <> 11 Then
            MsgBox "KeyError: cookie crumb assumption is invalid", vbCritical
            Exit Function
        End If
    End With

    ' call historical stock quotes from YF api
    ' construct api and get request
    api = "https://query1.finance.yahoo.com/v7/finance/download/" & stockTicker & "?period1=" & startDate & "&period2=" & endDate & "&interval=" & Frequency & "&events=history&crumb=" & crumb
    With objRequest_yf
        .Open "GET", api, False
        .setRequestHeader "Cookie", cookie
        .Send
        stockQuotes = .ResponseText
    End With
    GetStockQuote = stockQuotes

End Function
