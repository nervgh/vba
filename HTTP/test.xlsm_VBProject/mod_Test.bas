Attribute VB_Name = "mod_Test"
Sub Example()
    Dim NewHTML As Object
    Dim NewXML As Object
    Dim pageHTML As Object
    Dim pageXML As Object
    Dim pageText As String
    Dim Request As Object
    
    Const Url As String = "https://github.com/"
    
    With New HTTP
        '---------------------------------
        ' Create a HTML page
        Set NewHTML = .NewHTML
        '---------------------------------
        ' Create a XML page
        Set NewXML = .NewXML
        '---------------------------------
        ' Get a content of URL as a HTML
        Set pageHTML = .GetHTML(Url)
        '---------------------------------
        ' Get a content of URL as a XML
        Set pageXML = .GetXML(Url)
        '---------------------------------
        ' Get a content of URL as a text
        pageText = .GetText(Url)
        '--------------------------------
        ' Create a XMLHttpRequest
        Set Request = .Request
    End With
End Sub
