Attribute VB_Name = "modMath"
'=======================================================================
' Copyright 2005, Siebe "Inky" Tolsma, All rights reserved
'
' These functions were taken from Defyboy's code (All credit to him).
' I did however rewrite very large chunks of the functions to make them
' neater, faster and mostly better in handling data.
'=======================================================================
Option Explicit

Public Function hexXOr(strH1 As String, strH2 As String) As String
    Dim strB1 As String, strB2 As String
    strB1 = Hex2Bin(UCase$(strH1))
    strB2 = Hex2Bin(UCase$(strH2))
    
    Dim intLength As Integer: intLength = IIf(Len(strB1) > Len(strB2), Len(strB1), Len(strB2))
    strB1 = String$(intLength - Len(strB1), "0") & strB1
    strB2 = String$(intLength - Len(strB2), "0") & strB2
    
    Dim I As Integer, strXORed As String
    For I = 1 To intLength: strXORed = strXORed & (Mid$(strB1, I, 1) Xor Mid$(strB2, I, 1)): Next I
    hexXOr = strDecToHex(Bin2Hex(strXORed))
End Function

Public Function Hex2Bin(ByVal strHex As String) As String
    Dim strHexChar As String: strHexChar = "0123456789ABCDEF"
    Dim I As Integer, J As Integer, intDec As Integer
    
    For I = 1 To Len(strHex)
        intDec = InStr(1, strHexChar, Mid$(strHex, I, 1)) - 1
        For J = 3 To 0 Step -1
            Hex2Bin = Hex2Bin & CStr(intDec \ (2 ^ J))
            intDec = intDec Mod (2 ^ J)
        Next J
    Next I
    
    I = InStr(1, Hex2Bin, 1)
    If I <> 0 Then Hex2Bin = Mid$(Hex2Bin, I)
End Function

Public Function Bin2Hex(ByVal strBin As String) As String
    Dim I As Integer, intDec As Double
    
    strBin = String$(4 - (Len(strBin) Mod 4), "0") & strBin
    For I = 1 To Len(strBin)
        intDec = intDec + CInt(Mid$(strBin, Len(strBin) - I + 1, 1)) * 2 ^ (I - 1)
    Next I
    
    Bin2Hex = CStr(intDec)
    Bin2Hex = String$(Len(Bin2Hex) Mod 2, "0") & Bin2Hex
End Function

'=======================================================================
' These functions were taken from various sources. The authors remain
' unknown. I did a lot of optimizing again. If you think you are the
' author of one of these functions, please contact me so you
' can be credited :-)
'=======================================================================

Public Function strAdd(strA As String, strB As String) As String
    Dim intLenA As Integer, intLenB As Integer
    intLenA = Len(strA): intLenB = Len(strB)

    Dim intLongest As Integer: intLongest = IIf(intLenA > intLenB, intLenA, intLenB)
    strA = String$(intLongest - intLenA, "0") & strA
    strB = String$(intLongest - intLenB, "0") & strB
    
    Dim strOut As String, intTemp As Integer, intCarry As Integer, I As Integer
    strAdd = Space$(intLongest)
    For I = intLongest To 1 Step -1
        intTemp = Asc(Mid$(strA, I, 1)) + Asc(Mid$(strB, I, 1)) + intCarry - 96
        Mid$(strAdd, I, 1) = intTemp Mod 10
        intCarry = intTemp \ 10
    Next I
    
    If intCarry Then strAdd = "1" & strAdd
End Function

Public Function strMul(strA As String, strB As String) As String
    Dim intLenA As Integer, intLenB As Integer
    intLenA = Len(strA): intLenB = Len(strB)

    Dim intLongest As Integer: intLongest = IIf(intLenA > intLenB, intLenA, intLenB)
    strA = String$(intLongest - intLenA, "0") & strA
    strB = String$(intLongest - intLenB, "0") & strB
    
    Dim I As Integer, J As Integer, intCarry As Integer
    Dim strTemp As String, intTemp As Integer: strMul = "0"
    For I = intLongest To 1 Step -1
        intCarry = 0
    
        strTemp = Space$(intLongest + 1)
        For J = intLongest To 1 Step -1
            intTemp = ((Asc(Mid$(strA, I, 1)) - 48) * (Asc(Mid$(strB, J, 1)) - 48)) + intCarry
            
            Mid$(strTemp, J, 1) = intTemp Mod 10
            intCarry = intTemp \ 10
        Next J
        
        If intCarry Then strTemp = intCarry & strTemp
        strMul = strAdd(strMul, Trim$(strTemp) & String$(intLongest - I, "0"))
        DoEvents
    Next I
    
    strMul = Format$(strMul, "general number")
End Function

Public Function strMod(strA As String, strB As String) As String
    Dim intLenDif As Integer, I As Integer
    intLenDif = Len(strA) - Len(strB)
    If intLenDif < 0 Then intLenDif = 0

    strMod = strA
    For I = intLenDif To 0 Step -1
        While strGT(strMod, strB & String$(I, "0"))
            strMod = strSub(strMod, strB & String$(I, "0"))
        Wend
    Next I
    strMod = Format$(strMod, "general number")
End Function

Public Function strGT(strA As String, strB As String) As Boolean
    strA = Format$(strA, "general number")
    strB = Format$(strB, "general number")
    
    Dim intLenA As Integer, intLenB As Integer
    intLenA = Len(strA): intLenB = Len(strB)
    If intLenA > intLenB Then
        strGT = True
        Exit Function
    End If
    
    Dim intLongest As Integer: intLongest = IIf(intLenA > intLenB, intLenA, intLenB)
    strA = String$(intLongest - intLenA, "0") & strA
    strB = String$(intLongest - intLenB, "0") & strB
    
    Dim I As Integer, J As Integer, K As Integer: strGT = False
    For I = 1 To intLongest
        J = Asc(Mid$(strA, I, 1))
        K = Asc(Mid$(strB, I, 1))
        
        If J > K Then strGT = True
        If strGT Or K > J Then Exit For
    Next I
End Function

Private Function strSub(strA As String, strB As String) As String
    Dim intLenA As Integer, intLenB As Integer
    intLenA = Len(strA): intLenB = Len(strB)
    Dim intLongest As Integer: intLongest = IIf(intLenA > intLenB, intLenA, intLenB)
    strA = String$(intLongest - intLenA, "0") & strA
    strB = String$(intLongest - intLenB, "0") & strB
        
    strSub = Space$(intLongest)
    
    Dim I As Integer, J As Integer, K As Integer, intTemp As Integer
    For I = intLongest To 1 Step -1
        intTemp = Asc(Mid$(strA, I, 1)) - Asc(Mid$(strB, I, 1))
        
        If intTemp < 0 Then
            For J = I - 1 To 1 Step -1
                K = CInt(Mid$(strA, J, 1))
                If K > 0 Then Mid$(strA, J, 1) = CStr(K - 1): Exit For
            Next J
            
            If J = 0 Then strSub = "0": Exit Function
            
            For J = J + 1 To I: Mid$(strA, J, 1) = "9": Next J
            intTemp = intTemp + 10
        End If

        Mid$(strSub, I, 1) = intTemp
    Next I
End Function

Public Function strDecToHex(ByVal varDec As Variant) As String
    Dim intHexDigit As Double, dblDiv As Double
    strDecToHex = ""
    
    While varDec <> 0
        intHexDigit = varDec - (Int(varDec / 16) * 16)

        If intHexDigit < 10 Then
            strDecToHex = CStr(intHexDigit) & strDecToHex
        Else
            strDecToHex = Chr$(65 + intHexDigit - 10) & strDecToHex
        End If

        varDec = Int(varDec / 16)
    Wend

    If strDecToHex = "" Then strDecToHex = "0"
End Function

