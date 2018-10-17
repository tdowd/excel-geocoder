Attribute VB_Name = "geocoder"
Option Explicit

' Domain and URL for Google API
Public Const gstrGeocodingDomain = "https://maps.googleapis.com"
Public Const gstrGeocodingURL = "/maps/api/geocode/xml?"

' set gintType = 1 to use the Enterprise Geocoder (requires clientID and key)
' set gintType = 2 to use the API Premium Plan (requires key)
' leave gintType = 0 to use the free-ish Google geocoder (now requires a key! see https://developers.google.com/maps/documentation/geocoding/get-api-key)
Public Const gintType = 0

' key for Enterprise Geocoder or API Premium Plan or free-ish geocoder
Public Const gstrKey = ""

' clientID for Enterprise Geocoder
Public Const gstrClientID = ""

' kludge to not overdo the API calls and add a delay
#If VBA7 Then
    Public Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#Else
    Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#End If

Public Function AddressGeocode(address As String) As String
  Dim strAddress As String
  Dim strQuery As String
  Dim strLatitude As String
  Dim strLongitude As String
  Dim strQueryBland As String

  strAddress = URLEncode(address)

  'Assemble the query string
  strQuery = gstrGeocodingURL
  strQuery = strQuery & "address=" & strAddress
  If gintType = 0 Then ' free-ish Google Geocoder - now requires an API key!
    strQuery = strQuery & "&key=" & gstrKey
  ElseIf gintType = 1 Then ' Enterprise Geocoder
    strQuery = strQuery & "&client=" & gstrClientID
    strQuery = strQuery & "&signature=" & Base64_HMACSHA1(strQuery, gstrKey)
  ElseIf gintType = 2 Then ' API Premium Plan
    strQuery = strQuery & "&key=" & gstrKey
  End If

  'define XML and HTTP components
  Dim googleResult As New MSXML2.DOMDocument60
  Dim googleService As New MSXML2.XMLHTTP60
  Dim oNodes As MSXML2.IXMLDOMNodeList
  Dim oNode As MSXML2.IXMLDOMNode

  Sleep (5)

  'create HTTP request to query URL - make sure to have
  googleService.Open "GET", gstrGeocodingDomain & strQuery, False
  googleService.send
  googleResult.LoadXML (googleService.responseText)

  Set oNodes = googleResult.getElementsByTagName("geometry")

  If oNodes.Length = 1 Then
    For Each oNode In oNodes
      Debug.Print oNode.Text
      strLatitude = oNode.ChildNodes(0).ChildNodes(0).Text
      strLongitude = oNode.ChildNodes(0).ChildNodes(1).Text
      AddressGeocode = strLatitude & "," & strLongitude
    Next oNode
  Else
    AddressGeocode = "Not Found (try again, you may have done too many too fast)"
  End If
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


Public Function ReverseGeocode(lat As String, lng As String) As String
  Dim strAddress As String
  Dim strLat As String
  Dim strLng As String
  Dim strQuery As String
  Dim strLatitude As String
  Dim strLongitude As String

  strLat = URLEncode(lat)
  strLng = URLEncode(lng)

  'Assemble the query string
  strQuery = gstrGeocodingURL
  strQuery = strQuery & "latlng=" & strLat & "," & strLng
  If gintType = 0 Then ' free-ish Google Geocoder - now requires an API key!
    strQuery = strQuery & "&key=" & gstrKey
  ElseIf gintType = 1 Then ' Enterprise Geocoder
    strQuery = strQuery & "&client=" & gstrClientID
    strQuery = strQuery & "&signature=" & Base64_HMACSHA1(strQuery, gstrKey)
  ElseIf gintType = 2 Then ' API Premium Plan
    strQuery = strQuery & "&key=" & gstrKey
  End If

  'define XML and HTTP components
  Dim googleResult As New MSXML2.DOMDocument60
  Dim googleService As New MSXML2.XMLHTTP60
  Dim oNodes As MSXML2.IXMLDOMNodeList
  Dim oNode As MSXML2.IXMLDOMNode

  Sleep (5)

  'create HTTP request to query URL - make sure to have
  googleService.Open "GET", gstrGeocodingDomain & strQuery, False
  googleService.send
  googleResult.LoadXML (googleService.responseText)

  Set oNodes = googleResult.getElementsByTagName("formatted_address")
  
  If oNodes.Length > 0 Then
    ReverseGeocode = oNodes.Item(0).Text
  Else
    ReverseGeocode = "Not Found (try again, you may have done too many too fast)"
  End If
End Function


Public Function Base64_HMACSHA1(ByVal strTextToHash As String, ByVal strSharedSecretKey As String)

    Dim asc As Object
    Dim enc As Object
    Dim TextToHash() As Byte
    Dim SharedSecretKey() As Byte
    Dim bytes() As Byte
    
    Set asc = CreateObject("System.Text.UTF8Encoding")
    Set enc = CreateObject("System.Security.Cryptography.HMACSHA1")
    
    strSharedSecretKey = Replace(Replace(strSharedSecretKey, "-", "+"), "_", "/")
    SharedSecretKey = Base64Decode(strSharedSecretKey)
    enc.Key = SharedSecretKey
    
    TextToHash = asc.Getbytes_4(strTextToHash)
    bytes = enc.ComputeHash_2((TextToHash))
    Base64_HMACSHA1 = Replace(Replace(Base64Encode(bytes), "+", "-"), "/", "_")

End Function


Public Function Base64Decode(ByVal strData As String) As Byte()

Dim objXML As MSXML2.DOMDocument60
Dim objNode As MSXML2.IXMLDOMElement

Set objXML = New MSXML2.DOMDocument60
Set objNode = objXML.createElement("b64")
objNode.DataType = "bin.base64"
objNode.Text = strData
Base64Decode = objNode.nodeTypedValue

Set objNode = Nothing
Set objXML = Nothing

End Function


Public Function Base64Encode(ByRef arrData() As Byte) As String

Dim objXML As MSXML2.DOMDocument60
Dim objNode As MSXML2.IXMLDOMElement

Set objXML = New MSXML2.DOMDocument60
Set objNode = objXML.createElement("b64")
objNode.DataType = "bin.base64"
objNode.nodeTypedValue = arrData
Base64Encode = objNode.Text

Set objNode = Nothing
Set objXML = Nothing

End Function
