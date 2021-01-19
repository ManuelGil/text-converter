Attribute VB_Name = "ANSICoder"
Option Explicit

' This method converts the hexadecimal number to decimal equivalent.

Public Function HEXToDEC(ByVal s As String) As String
    Dim i As Integer
    Dim sArray() As String
    Dim output As String
    
    ' Cut a text and save them in an array
    sArray = Split(s, " ")
    
    ' Loop through the each items
    For i = LBound(sArray) To UBound(sArray)
        ' Convert the item to decimal
        output = output + Str(Dec(UCase(sArray(i)))) + " "
    Next i
    
    ' Return the new string
    HEXToDEC = Trim(output)
End Function

' This method converts the hexadecimal number to decimal equivalent.

Public Function DECToHEX(ByVal s As String) As String
    Dim i As Integer
    Dim sArray() As String
    Dim output As String
    
    ' Cut a text and save them in an array.
    sArray = Split(s, " ")
    
    ' Loop through the each items.
    For i = LBound(sArray) To UBound(sArray)
        ' Convert the item to hexadecimal.
        output = output + Hex(sArray(i)) + " "
    Next i
    
    ' Return the new string.
    DECToHEX = Trim(output)
End Function

' This method transforms a character set to hexadecimal equivalent.

Public Function HEXEncode(ByVal s As String) As String
    Dim i As Integer
    Dim output As String
    
    ' Loop through the each character.
    For i = 1 To Len(s)
        ' Transform the character to hexadecimal.
        output = output + Hex(Asc(Mid(s, i, 1))) + " "
    Next i
    
    ' Return the new string.
    HEXEncode = Trim(output)
End Function

' This method transforms a hexadecimal set to character set equivalent.

Public Function HEXDecode(ByVal s As String) As String
    Dim i As Integer
    Dim sArray() As String
    Dim output As String
    
    ' Cut a text and save them in an array.
    sArray = Split(s, " ")
    
    ' Loop through the each item.
    For i = LBound(sArray) To UBound(sArray)
        ' Transform the hexadecimal to character.
        output = output + Chr(Dec(UCase(sArray(i))))
    Next i
    
    ' Return the new string.
    HEXDecode = output
End Function

' This method transforms a character set to decimal set equivalent.

Public Function DECEncode(ByVal s As String) As String
    Dim i As Integer
    Dim output As String
    
    ' Loop through the each item.
    For i = 1 To Len(s)
        ' Transform the character to decimal.
        output = output + Str(Asc(Mid(s, i, 1))) + " "
    Next i
    
    ' Return the new string.
    DECEncode = Trim(output)
End Function

' This method transforms a decimal set to character set equivalent.

Public Function DECDecode(ByVal s As String) As String
    Dim i As Integer
    Dim sArray() As String
    Dim output As String
    
    ' Loop through the each item.
    sArray = Split(s, " ")
    
    For i = LBound(sArray) To UBound(sArray)
        ' Transform the decimal to character.
        output = output + Chr(UCase(sArray(i)))
    Next i
    
    ' Return the new string.
    DECDecode = output
End Function

' This method converts hexadecimal to decimal.

Private Function Dec(ByVal Hexa As String) As Variant
    ' Return the conversion.
    Dec = CDec("&H" & Hexa)
End Function

' This method transforms a character set to Visual Basic notation equivalent.

Public Function VBEncode(ByVal s As String) As String
    Dim i As Integer
    Dim output As String
    
    ' Loop through the each item.
    For i = 1 To Len(s) - 1
        ' Transform the character to hexadecimal and add format.
        output = output + "Chr(&H" + Hex(Asc(Mid(s, i, 1))) + ") + "
    Next i
    
    ' Transform the last character to hexadecimal and add format.
    output = output + "Chr(&H" + Hex(Asc(Mid(s, i, 1))) + ")"
    
    ' Return the new string.
    VBEncode = output
End Function

' This method transforms a character set to Java notation equivalent.

Public Function JAVAEncode(ByVal s As String) As String
    Dim i As Integer
    Dim output As String
    
    ' Loop through the each item.
    For i = 1 To Len(s) - 1
        ' Transform the character to hexadecimal and add format.
        output = output + "\u00" + Hex(Asc(Mid(s, i, 1))) + " + "
    Next i
    
    ' Transform the last character to hexadecimal and add format.
    output = output + "\u00" + Hex(Asc(Mid(s, i, 1)))
    
    ' Return the new string.
    JAVAEncode = output
End Function

