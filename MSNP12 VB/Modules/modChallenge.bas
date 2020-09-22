Attribute VB_Name = "modChallenge"
Option Explicit

'=======================================================================
' Copyright 2005, Siebe "Inky" Tolsma, All rights reserved
' Math functions based on Troy Osborne's (admin@defyboy.com) implementation
' Initially based on documentation by ZoRoNaX
'=======================================================================

Public Function CreateQRY(strChallenge As String, Optional strClientID As String = "PROD0090YUAUV{2B", Optional strClientCode As String = "YMM8C_H7KCQ2S_KL")
    'First we need to create 32 bit integers from an MD5 Hash
    Dim strMD5 As String, strMD5Ints() As String
    strMD5 = MD5_Hex(strChallenge & strClientCode)
    strMD5Ints = MD5HexToInt(strMD5)
 
    'Then we repeat the process with almost the same steps, just with two different strings
    Dim strCHL As String, strCHLInts() As String
    strCHL = strChallenge & strClientID
    strCHL = strCHL & String$(8 - (Len(strCHL) Mod 8), "0")
    strCHLInts = CHLHexToInt(strCHL)

    'Create the XOR key (Hi/Lo) :-)
    Dim strXORKey As String, strHigh As String, strLow As String
    strXORKey = CreateKey(strMD5Ints, strCHLInts)
    strHigh = Mid$(strXORKey, 1, 8): strLow = Mid$(strXORKey, 9, 8)

    'And finally XOR the key
    Dim strMD5P() As String: strMD5P = ChopString(strMD5, 8)
    strMD5P(0) = LCase$(hexXOr(strMD5P(0), strHigh) & hexXOr(strMD5P(1), strLow))
    strMD5P(1) = LCase$(hexXOr(strMD5P(2), strHigh) & hexXOr(strMD5P(3), strLow))
    
    'Pad it and return it :-)
    CreateQRY = String$(16 - Len(strMD5P(0)), "0") & strMD5P(0) & _
                String$(16 - Len(strMD5P(1)), "0") & strMD5P(1)
End Function

Private Function MD5HexToInt(strMD5 As String) As String()
    'Chop the MD5 hash into pieces of 8
    Dim strMD5Ints() As String, I As Integer
    strMD5Ints = ChopString(strMD5, 8)
    
    'Loop over the chunks given to use and create the appropriate integers from it
    For I = 0 To UBound(strMD5Ints)
        'Store the value :-)
        strMD5Ints(I) = CStr(CDbl("&H" & SwapBytes(strMD5Ints(I))) And &H7FFFFFFF)
    Next I
    
    'Return them
    MD5HexToInt = strMD5Ints
End Function

Private Function CHLHexToInt(strCHL As String) As String()
    'Chop the string into pieces of 4
    Dim strCHLInts() As String, I As Integer
    strCHLInts = ChopString(strCHL, 4)
    
    'Loop over the entries in the array and create integers from them
    For I = 0 To UBound(strCHLInts)
        'Store the value :-)
        strCHLInts(I) = CStr(CDbl("&H" & SwapBytes(BinToHex(strCHLInts(I)))))
    Next I
    
    'Return them
    CHLHexToInt = strCHLInts
End Function

Private Function CreateKey(strMD5Ints() As String, strCHLInts() As String, Optional strMagicKey As String = "&H0E79A9C1") As String
    'Initialize variables |-)
    Dim strHigh As String, strLow As String, strTemp As String
    strHigh = "0": strLow = "0": strTemp = "0"
    
    'And some more (So we dont have to calculate these each time
    Dim strH7F As String, I As Integer
    strH7F = CStr(Int("&H7FFFFFFF"))
    strMagicKey = CStr(Int(strMagicKey))
    
    'Then walk over the strCHLInts array (Stepping 2 at a time)
    For I = 0 To UBound(strCHLInts) Step 2
        'First calculate the temporary variable
        strTemp = strMod(strMul(strCHLInts(I), strMagicKey), strH7F)
        strTemp = strMul(strAdd(strTemp, strHigh), strMD5Ints(0))
        strTemp = strMod(strAdd(strTemp, strMD5Ints(1)), strH7F)
        
        'Then the high part of the key
        strHigh = strMod(strAdd(strCHLInts(I + 1), strTemp), strH7F)
        strHigh = strAdd(strMul(strHigh, strMD5Ints(2)), strMD5Ints(3))
        strHigh = strMod(strHigh, strH7F)
        
        'Then add them to the low part of the key
        strLow = strAdd(strAdd(strLow, strHigh), strTemp)
    Next I
    
    'Final step of the official part :-)
    strHigh = strMod(strAdd(strHigh, strMD5Ints(1)), strH7F)
    strLow = strMod(strAdd(strLow, strMD5Ints(3)), strH7F)
    
    'Swap the bytes around and output as hex
    CreateKey = SwapBytes(strDecToHex(strHigh)) & SwapBytes(strDecToHex(strLow))
End Function

Private Function ChopString(strString As String, iLength As Integer) As String()
    Dim strChunks() As String, I As Integer
    
    'Create a For loop
    For I = 0 To Len(strString) - 1 Step iLength
        'Redim the array accordingly, "Push" the value into the array
        ReDim Preserve strChunks(I / iLength) As String
        strChunks(I / iLength) = Mid$(strString, I + 1, iLength)
    Next I
    
    'Pass it back to wherever
    ChopString = strChunks
End Function

Private Function SwapBytes(strValue As String) As String
    'Swap the bytes around for this value (Hex = No overflow ^_^)
    Dim I As Integer
    For I = 1 To Len(strValue) Step 2
        'Take each 2 characters and put them up front, slowly swapping bytes
        SwapBytes = Mid$(strValue, I, 2) & SwapBytes
    Next I
End Function

Private Function BinToHex(strString As String) As String
    'Output the string as hex
    Dim I As Integer
    For I = 1 To Len(strString)
        'Take a character, find the ASCII value and convert it to Hex (Base 16)
        BinToHex = BinToHex & Hex$(Asc(Mid$(strString, I, 1)))
    Next I
End Function
