VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "HTTP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' @author nerv
' @version 18/01/2013, 0.2


' ----------------------------------------
' Returns a new HTML object
'
' @return {Object} NewHTML
' ----------------------------------------
Public Property Get NewHTML() As Object
    Set NewHTML = CreateObject("HTMLFile")
End Property


' ----------------------------------------
' Returns a new XML object
'
' @return {Object} NewXML
' ----------------------------------------
Public Property Get NewXML() As Object
    Set NewXML = CreateObject("MSXML.DOMDocument")
End Property


' ----------------------------------------
' Returns a new XMLHTTPRequest object
'
' @return {Object} Request
' ----------------------------------------
Public Property Get Request() As Object
    Set Request = CreateObject("MSXML2.XMLHTTP")
End Property


' ----------------------------------------
' Returns a content of the URL as as a text
'
' @param {String} Url
' @return {String} GetText
' ----------------------------------------
Public Property Get GetText(ByVal Url As String) As String
    GetText = DefaultRequest(Url).responseText
End Property


' ----------------------------------------
' Returns a content of the URL as a HTML page
'
' @param {String} Url
' @return {Object} GetHTML
' ----------------------------------------
Public Property Get GetHTML(ByVal Url As String) As Object
    Set GetHTML = Me.NewHTML
    GetHTML.Write Me.GetText(Url)
End Property


' ----------------------------------------
' Returns a content of the URL as a XML page
'
' @param {String} Url
' @return {Object} GetXML
' ----------------------------------------
Public Property Get GetXML(ByVal Url As String) As Object
    Set GetXML = DefaultRequest(Url).responseXML
End Property


' ----------------------------------------
' Sends a synchronous request and returns a XMLHTTPRequest object
'
' @param {String} Url
' @return {Object} DefaultRequest
' ----------------------------------------
Private Function DefaultRequest(ByVal Url As String) As Object
    Set DefaultRequest = Me.Request
    
    With DefaultRequest
        .Open "GET", Url, False
        
        ' no internet connection
        On Error Resume Next
            .Send Null
        On Error GoTo 0
    End With
End Function
