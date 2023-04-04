Attribute VB_Name = "ModBinanceRequests"
' function takes a string to hash, a secret key to hash with and returns a signature for Binance HTTP requests
Function getSignature(toHash As String, secretKey As String)
    getSignature = Hex_HMACSHA256(toHash, secretKey)
End Function

' function takes a string to hash, a secret key to hash with HMAC SHA256 algorithm, and returns an hexadecimal string
' inspired from internet
Public Function Hex_HMACSHA256(ByVal sTextToHash As String, ByVal sSharedSecretKey As String)
    Dim asc As Object, enc As Object
    Dim TextToHash() As Byte
    Dim SharedSecretKey() As Byte
    Set asc = CreateObject("System.Text.UTF8Encoding")
    Set enc = CreateObject("System.Security.Cryptography.HMACSHA256")
    TextToHash = asc.Getbytes_4(sTextToHash)
    SharedSecretKey = asc.Getbytes_4(sSharedSecretKey)
    enc.key = SharedSecretKey
    Dim bytes() As Byte
    bytes = enc.ComputeHash_2((TextToHash))
    ' tried to encode bytes directly to HEX but not working so have to do bytes > b64 > b16
    Hex_HMACSHA256 = LCase(Base64To16(EncodeBase64(bytes)))
    Set asc = Nothing
    Set enc = Nothing
End Function

' function takes an array of bytes and returns b64 encoded string
' inspired from internet
Private Function EncodeBase64(ByRef arrData() As Byte) As String
    'Inside the VBE, Go to Tools -> References, then Select Microsoft XML, v6.0
    '(or whatever your latest is. This will give you access to the XML Object Library.)
    Dim objXML As MSXML2.DOMDocument60
    Dim objNode As MSXML2.IXMLDOMElement
    Set objXML = New MSXML2.DOMDocument60
    ' byte array to base64
    Set objNode = objXML.createElement("b64")
    objNode.DataType = "bin.base64"
    objNode.nodeTypedValue = arrData
    EncodeBase64 = objNode.Text
    Set objNode = Nothing
    Set objXML = Nothing
End Function

' function to directly encode bytes into HEX
' I think it works alone but not with the HEX HMAC SHA 256 function
' inspired from internet
Private Function ByteArrayToHex(ByRef ByteArray() As Byte) As String
    Dim l As Long, strRet As String
    For l = LBound(ByteArray) To UBound(ByteArray)
        strRet = strRet & Hex$(ByteArray(l)) & ""
    Next l
    ByteArrayToHex = LCase(strRet)
End Function

' function returns the current unix timestamp of the binance server
Function getTimeStampForBinance()
'    Tout ça n'a servi à rien, je ne savais pas
'    qu 'on pouvait directement avoir le server time
'    de binance

'    Dim timestamp_string As String
'    Dim timestamp_vba As Double
'    Dim timestamp_real As Double
'    Dim timestamp_binance As Double
'
'    timestamp_string = Split((DateDiff("s", "01/01/1970", Date) + Timer) * 1000, ",")(0)
'    'MsgBox timestamp_string
'    timestamp_vba = CDbl(timestamp_string)
'    ' VBA's ts is 3594104 ahead
'    timestamp_real = timestamp_vba - 3594104
'    timestamp_binance = timestamp_real - 70000 '50000 sur les PC de la FAC
    Dim xmlhttp As Object
    Set xmlhttp = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    Dim json As Object
    url = "https://testnet.binance.vision/api/v3/exchangeInfo"
    xmlhttp.Open "GET", url, False
    xmlhttp.Send
    Set json = JsonConverter.ParseJson(xmlhttp.responseText)
    getTimeStampForBinance = json("serverTime")
End Function

' function takes b64 string and returns b16 string
' inspired from internet
Function Base64To16(Base64 As String) As String
  Dim Base2 As String
  Dim i As Long
  If Len(Base64) Mod 4 > 0 Then
    Base64To16 = CVErr(xlErrValue)
    Exit Function
  End If
  For i = 1 To Len(Base64)
    Select Case Mid(Base64, i, 1)
      Case "A"
        Base2 = Base2 & "000000"
      Case "B"
        Base2 = Base2 & "000001"
      Case "C"
        Base2 = Base2 & "000010"
      Case "D"
        Base2 = Base2 & "000011"
      Case "E"
        Base2 = Base2 & "000100"
      Case "F"
        Base2 = Base2 & "000101"
      Case "G"
        Base2 = Base2 & "000110"
      Case "H"
        Base2 = Base2 & "000111"
      Case "I"
        Base2 = Base2 & "001000"
      Case "J"
        Base2 = Base2 & "001001"
      Case "K"
        Base2 = Base2 & "001010"
      Case "L"
        Base2 = Base2 & "001011"
      Case "M"
        Base2 = Base2 & "001100"
      Case "N"
        Base2 = Base2 & "001101"
      Case "O"
        Base2 = Base2 & "001110"
      Case "P"
        Base2 = Base2 & "001111"
      Case "Q"
        Base2 = Base2 & "010000"
      Case "R"
        Base2 = Base2 & "010001"
      Case "S"
        Base2 = Base2 & "010010"
      Case "T"
        Base2 = Base2 & "010011"
      Case "U"
        Base2 = Base2 & "010100"
      Case "V"
        Base2 = Base2 & "010101"
      Case "W"
        Base2 = Base2 & "010110"
      Case "X"
        Base2 = Base2 & "010111"
      Case "Y"
        Base2 = Base2 & "011000"
      Case "Z"
        Base2 = Base2 & "011001"
      Case "a"
        Base2 = Base2 & "011010"
      Case "b"
        Base2 = Base2 & "011011"
      Case "c"
        Base2 = Base2 & "011100"
      Case "d"
        Base2 = Base2 & "011101"
      Case "e"
        Base2 = Base2 & "011110"
      Case "f"
        Base2 = Base2 & "011111"
      Case "g"
        Base2 = Base2 & "100000"
      Case "h"
        Base2 = Base2 & "100001"
      Case "i"
        Base2 = Base2 & "100010"
      Case "j"
        Base2 = Base2 & "100011"
      Case "k"
        Base2 = Base2 & "100100"
      Case "l"
        Base2 = Base2 & "100101"
      Case "m"
        Base2 = Base2 & "100110"
      Case "n"
        Base2 = Base2 & "100111"
      Case "o"
        Base2 = Base2 & "101000"
      Case "p"
        Base2 = Base2 & "101001"
      Case "q"
        Base2 = Base2 & "101010"
      Case "r"
        Base2 = Base2 & "101011"
      Case "s"
        Base2 = Base2 & "101100"
      Case "t"
        Base2 = Base2 & "101101"
      Case "u"
        Base2 = Base2 & "101110"
      Case "v"
        Base2 = Base2 & "101111"
      Case "w"
        Base2 = Base2 & "110000"
      Case "x"
        Base2 = Base2 & "110001"
      Case "y"
        Base2 = Base2 & "110010"
      Case "z"
        Base2 = Base2 & "110011"
      Case "0"
        Base2 = Base2 & "110100"
      Case "1"
        Base2 = Base2 & "110101"
      Case "2"
        Base2 = Base2 & "110110"
      Case "3"
        Base2 = Base2 & "110111"
      Case "4"
        Base2 = Base2 & "111000"
      Case "5"
        Base2 = Base2 & "111001"
      Case "6"
        Base2 = Base2 & "111010"
      Case "7"
        Base2 = Base2 & "111011"
      Case "8"
        Base2 = Base2 & "111100"
      Case "9"
        Base2 = Base2 & "111101"
      Case "+"
        Base2 = Base2 & "111110"
      Case "/"
        Base2 = Base2 & "111111"
      Case "="
        Base2 = Left(Base2, Len(Base2) - 2)
      Case Else
        Base64To16 = CVErr(xlErrValue)
        Exit Function
    End Select
  Next i
  If Not Len(Base2) Mod 4 = 0 Then
    Base2 = String(4 - (Len(Base2) Mod 4), "0") & Base2
  End If
  If Len(Base2) > 4 And Left(Base2, 4) = "0000" Then
    Base2 = Mid(Base2, 5)
  End If
  Base64To16 = ""
  For i = 1 To Len(Base2) Step 4
    Select Case Mid(Base2, i, 4)
      Case "0000"
        Base64To16 = Base64To16 & "0"
      Case "0001"
        Base64To16 = Base64To16 & "1"
      Case "0010"
        Base64To16 = Base64To16 & "2"
      Case "0011"
        Base64To16 = Base64To16 & "3"
      Case "0100"
        Base64To16 = Base64To16 & "4"
      Case "0101"
        Base64To16 = Base64To16 & "5"
      Case "0110"
        Base64To16 = Base64To16 & "6"
      Case "0111"
        Base64To16 = Base64To16 & "7"
      Case "1000"
        Base64To16 = Base64To16 & "8"
      Case "1001"
        Base64To16 = Base64To16 & "9"
      Case "1010"
        Base64To16 = Base64To16 & "A"
      Case "1011"
        Base64To16 = Base64To16 & "B"
      Case "1100"
        Base64To16 = Base64To16 & "C"
      Case "1101"
        Base64To16 = Base64To16 & "D"
      Case "1110"
        Base64To16 = Base64To16 & "E"
      Case "1111"
        Base64To16 = Base64To16 & "F"
    End Select
  Next i
  If Len(Base64To16) Mod 2 = 1 Then
    Base64To16 = "0" & Base64To16
  End If
End Function
