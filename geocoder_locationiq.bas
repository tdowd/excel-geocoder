Attribute VB_Name = "geocoder"
Option Explicit

Public Const gstrKey = ""
' request a key from http://locationiq.org
Public Const gstrGeocodingDomain = "http://locationiq.org/v1/search.php?key=" & gstrKey & "&format=xml&q="

Function AddressGeocode(address As String) As String
  Dim strAddress As String
  Dim strQuery As String
  Dim strLatitude As String
  Dim strLongitude As String
  Dim strQueryBland As String

  strAddress = URLEncode(address)

  'Assemble the query string
  strQuery = gstrGeocodingDomain
  strQuery = strQuery & strAddress

  'define XML and HTTP components
  Dim XMLResult As New MSXML2.DOMDocument
  Dim XMLService As New MSXML2.XMLHTTP
  Dim oNodes As MSXML2.IXMLDOMNodeList
  Dim oNode As MSXML2.IXMLDOMNode

  'create HTTP request to query URL - make sure to have
  XMLService.Open "GET", gstrGeocodingDomain & strQuery, False
  XMLService.send
  XMLResult.LoadXML (XMLService.responseText)

  Set oNodes = XMLResult.getElementsByTagName("place")

  Dim i As Integer

  For i = 0 To i = 1
      AddressGeocode = oNodes(i).Attributes.getNamedItem("lat").Text & "," & oNodes(i).Attributes.getNamedItem("lon").Text
  Next
  
End Function


Public Function URLEncode(StringToEncode As String, Optional _
   UsePlusRatherThanHexForSpace As Boolean = False) As String

Dim TempAns As String
Dim CurChr As Integer
CurChr = 1
Do Until CurChr - 1 = Len(StringToEncode)
  Select Case asc(Mid(StringToEncode, CurChr, 1))
    Case 48 To 57, 65 To 90, 97 To 122
      TempAns = TempAns & Mid(StringToEncode, CurChr, 1)
    Case 32
      If UsePlusRatherThanHexForSpace = True Then
        TempAns = TempAns & "+"
      Else
        TempAns = TempAns & "%" & Hex(32)
      End If
   Case Else
         TempAns = TempAns & "%" & _
              Format(Hex(asc(Mid(StringToEncode, _
              CurChr, 1))), "00")
End Select

  CurChr = CurChr + 1
Loop

URLEncode = TempAns
End Function
