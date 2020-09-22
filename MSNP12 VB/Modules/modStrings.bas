Attribute VB_Name = "modStrings"
' --- Function to Send Data to the NS Server
Public Function MSNDecode(ByVal Utf8Str As String) As String
    Utf8Str = Replace(Utf8Str, "%20", " ")
    Utf8Str = Replace(Utf8Str, "ツ", "?")
    Utf8Str = Replace(Utf8Str, "™", "�")
    Utf8Str = Replace(Utf8Str, "�&#8218;�", "&#8364;")
    Utf8Str = Replace(Utf8Str, "", "�")
    Utf8Str = Replace(Utf8Str, "�&#8364;&#353;", "&#8218;")
    Utf8Str = Replace(Utf8Str, "�&#8217;", "&#402;")
    Utf8Str = Replace(Utf8Str, "�&#8364;&#382;", "&#8222;")
    Utf8Str = Replace(Utf8Str, "�&#8364;�", "&#8230;")
    Utf8Str = Replace(Utf8Str, "�&#8364; ", "&#8224;")
    Utf8Str = Replace(Utf8Str, "�&#8364;�", "&#8225;")
    Utf8Str = Replace(Utf8Str, "�&#8224;", "&#710;")
    Utf8Str = Replace(Utf8Str, "�&#8364;�", "&#8240;")
    Utf8Str = Replace(Utf8Str, "� ", "&#352;")
    Utf8Str = Replace(Utf8Str, "�&#8364;�", "&#8249;")
    Utf8Str = Replace(Utf8Str, "�&#8217;", "&#338;")
    Utf8Str = Replace(Utf8Str, "", "�")
    Utf8Str = Replace(Utf8Str, "Ž", "&#381;")
    Utf8Str = Replace(Utf8Str, "", "�")
    Utf8Str = Replace(Utf8Str, "", "�")
    Utf8Str = Replace(Utf8Str, "�&#8364;&#732;", "&#8216;")
    Utf8Str = Replace(Utf8Str, "�&#8364;&#8482;", "&#8217;")
    Utf8Str = Replace(Utf8Str, "�&#8364;&#339;", "&#8220;")
    Utf8Str = Replace(Utf8Str, "�&#8364;�", "&#8221;")
    Utf8Str = Replace(Utf8Str, "�&#8364;�", "&#8226;")
    Utf8Str = Replace(Utf8Str, "�&#8364;&#8220;", "&#8211;")
    Utf8Str = Replace(Utf8Str, "�&#8364;&#8221;", "&#8212;")
    Utf8Str = Replace(Utf8Str, "�&#339;", "&#732;")
    Utf8Str = Replace(Utf8Str, "�&#8222;�", "&#8482;")
    Utf8Str = Replace(Utf8Str, "š", "&#353;")
    Utf8Str = Replace(Utf8Str, "�&#8364;�", "&#8250;")
    Utf8Str = Replace(Utf8Str, "�&#8220;", "&#339;")
    Utf8Str = Replace(Utf8Str, "", "�")
    Utf8Str = Replace(Utf8Str, "'ž", "&#382;")
    Utf8Str = Replace(Utf8Str, "Ÿ", "&#376;")
    Utf8Str = Replace(Utf8Str, "� ", " ")
    Utf8Str = Replace(Utf8Str, "¡", "�")
    Utf8Str = Replace(Utf8Str, "¢", "�")
    Utf8Str = Replace(Utf8Str, "£", "�")
    Utf8Str = Replace(Utf8Str, "¤", "�")
    Utf8Str = Replace(Utf8Str, "¥", "�")
    Utf8Str = Replace(Utf8Str, "¦", "�")
    Utf8Str = Replace(Utf8Str, "§", "�")
    Utf8Str = Replace(Utf8Str, "¨", "�")
    Utf8Str = Replace(Utf8Str, "©", "�")
    Utf8Str = Replace(Utf8Str, "ª", "�")
    Utf8Str = Replace(Utf8Str, "«", "�")
    Utf8Str = Replace(Utf8Str, "¬", "�")
    Utf8Str = Replace(Utf8Str, "­", "�")
    Utf8Str = Replace(Utf8Str, "®", "�")
    Utf8Str = Replace(Utf8Str, "¯", "�")
    Utf8Str = Replace(Utf8Str, "°", "�")
    Utf8Str = Replace(Utf8Str, "±", "�")
    Utf8Str = Replace(Utf8Str, "²", "�")
    Utf8Str = Replace(Utf8Str, "³", "�")
    Utf8Str = Replace(Utf8Str, "´", "�")
    Utf8Str = Replace(Utf8Str, "µ", "�")
    Utf8Str = Replace(Utf8Str, "¶", "�")
    Utf8Str = Replace(Utf8Str, "·", "�")
    Utf8Str = Replace(Utf8Str, "¸", "�")
    Utf8Str = Replace(Utf8Str, "¹", "�")
    Utf8Str = Replace(Utf8Str, "º", "�")
    Utf8Str = Replace(Utf8Str, "»", "�")
    Utf8Str = Replace(Utf8Str, "¼", "�")
    Utf8Str = Replace(Utf8Str, "½", "�")
    Utf8Str = Replace(Utf8Str, "¾", "�")
    Utf8Str = Replace(Utf8Str, "¿", "�")
    Utf8Str = Replace(Utf8Str, "� ", "�")
    Utf8Str = Replace(Utf8Str, "á", "�")
    Utf8Str = Replace(Utf8Str, "â", "�")
    Utf8Str = Replace(Utf8Str, "ã", "�")
    Utf8Str = Replace(Utf8Str, "ä", "�")
    Utf8Str = Replace(Utf8Str, "å", "�")
    Utf8Str = Replace(Utf8Str, "æ", "�")
    Utf8Str = Replace(Utf8Str, "ç", "�")
    Utf8Str = Replace(Utf8Str, "è", "�")
    Utf8Str = Replace(Utf8Str, "é", "�")
    Utf8Str = Replace(Utf8Str, "ê", "�")
    Utf8Str = Replace(Utf8Str, "ë", "�")
    Utf8Str = Replace(Utf8Str, "ì", "�")
    Utf8Str = Replace(Utf8Str, "í", "�")
    Utf8Str = Replace(Utf8Str, "î", "�")
    Utf8Str = Replace(Utf8Str, "ï", "�")
    Utf8Str = Replace(Utf8Str, "ð", "�")
    Utf8Str = Replace(Utf8Str, "ñ", "�")
    Utf8Str = Replace(Utf8Str, "ò", "�")
    Utf8Str = Replace(Utf8Str, "ó", "�")
    Utf8Str = Replace(Utf8Str, "ô", "�")
    Utf8Str = Replace(Utf8Str, "õ", "�")
    Utf8Str = Replace(Utf8Str, "ö", "�")
    Utf8Str = Replace(Utf8Str, "÷", "�")
    Utf8Str = Replace(Utf8Str, "ø", "�")
    Utf8Str = Replace(Utf8Str, "ù", "�")
    Utf8Str = Replace(Utf8Str, "ú", "�")
    Utf8Str = Replace(Utf8Str, "û", "�")
    Utf8Str = Replace(Utf8Str, "ü", "�")
    Utf8Str = Replace(Utf8Str, "ý", "�")
    Utf8Str = Replace(Utf8Str, "þ", "�")
    Utf8Str = Replace(Utf8Str, "ÿ", "�")
    Utf8Str = Replace(Utf8Str, "�&#8364;", "�")
    Utf8Str = Replace(Utf8Str, "Á", "�")
    Utf8Str = Replace(Utf8Str, "�&#8218;", "�")
    Utf8Str = Replace(Utf8Str, "�&#402;", "�")
    Utf8Str = Replace(Utf8Str, "�&#8222;", "�")
    Utf8Str = Replace(Utf8Str, "�&#8230;", "�")
    Utf8Str = Replace(Utf8Str, "�&#8224;", "�")
    Utf8Str = Replace(Utf8Str, "�&#8225;", "�")
    Utf8Str = Replace(Utf8Str, "�&#710;", "�")
    Utf8Str = Replace(Utf8Str, "�&#8240;", "�")
    Utf8Str = Replace(Utf8Str, "�&#352;", "�")
    Utf8Str = Replace(Utf8Str, "�&#8249;", "�")
    Utf8Str = Replace(Utf8Str, "�&#338;", "�")
    Utf8Str = Replace(Utf8Str, "Í", "�")
    Utf8Str = Replace(Utf8Str, "�&#381;", "�")
    Utf8Str = Replace(Utf8Str, "Ï", "�")
    Utf8Str = Replace(Utf8Str, "Ð", "�")
    Utf8Str = Replace(Utf8Str, "�&#8216;", "�")
    Utf8Str = Replace(Utf8Str, "�&#8217;", "�")
    Utf8Str = Replace(Utf8Str, "�&#8220;", "�")
    Utf8Str = Replace(Utf8Str, "�&#8221;", "�")
    Utf8Str = Replace(Utf8Str, "�&#8226;", "�")
    Utf8Str = Replace(Utf8Str, "�&#8211;", "�")
    Utf8Str = Replace(Utf8Str, "�&#8212;", "�")
    Utf8Str = Replace(Utf8Str, "�&#732;", "�")
    Utf8Str = Replace(Utf8Str, "�&#8482;", "�")
    Utf8Str = Replace(Utf8Str, "�&#353;", "�")
    Utf8Str = Replace(Utf8Str, "�&#8250;", "�")
    Utf8Str = Replace(Utf8Str, "�&#339;", "�")
    Utf8Str = Replace(Utf8Str, "Ý", "�")
    Utf8Str = Replace(Utf8Str, "�&#382;", "�")
    Utf8Str = Replace(Utf8Str, "�&#376;", "�")
    Utf8Str = Replace(Utf8Str, "%40", "@")
    Utf8Str = Replace(Utf8Str, "%2E", ".")
    Utf8Str = Replace(Utf8Str, "%20", " ")
    MSNDecode = Utf8Str
End Function

Public Function URLEncode(s As String) As String
    Dim sChar As String, sAsc As String, sHex As String, sName As String
    Dim I As Long
    
    'Encode this string to URL format
    For I = 1 To Len(s)
        sChar = Mid$(s, I, 1)
        sAsc = Asc(sChar)

        If (sAsc > 44 And sAsc < 59) _
        Or (sAsc > 64 And sAsc < 94) _
        Or (sAsc > 96 And sAsc < 126) Then
            sHex = sChar
        Else
            sHex = "%" & Hex(sAsc)
        End If
        
        sName = sName & sHex
    Next I

    URLEncode = Replace$(sName, "%D%A", "%0D")
End Function

Public Function URLDecode(s As String) As String
    Dim sChar As String, sHex As String, sName As String
    Dim I As Long
    
    'Get the Unicode name
    If InStr(1, s, "%") Then
        For I = 1 To Len(s)
            sChar = Mid$(s, I, 1)
            sHex = Mid$(s, I + 1, 2)
    
            If sChar = "%" Then
                sName = sName & Chr$(Val("&H" & sHex)): I = I + 2
            Else
                sName = sName & sChar
            End If
        Next I
    Else
        sName = s
    End If
    URLDecode = sName
End Function


Public Function GetBetween(Str As String, Optional dStart As String, Optional dEnd As String, Optional Length As Long) As String
    Dim x1 As Long, x2 As Long

    x1 = IIf(dStart = "", 1, InStr(1, LCase$(Str), LCase$(dStart)) + Len(dStart))
    If x1 > 0 Then
        If dEnd = "" Then
            GetBetween = Mid$(Str, x1)
        Else
            x2 = InStr(x1, LCase$(Str), LCase$(dEnd)) - x1
            If x2 > 0 Then
                GetBetween = Mid$(Str, x1, x2)
            Else
                GetBetween = "n/f"
            End If
        End If
    Else
        GetBetween = "n/f"
    End If
    If Length > 0 And GetBetween <> "n/f" Then GetBetween = Left$(GetBetween, Length)
End Function

Public Function SendData(Data As String)
    Data = Replace(Data, "#", frmMain.TriID)
    Call frmMain.sckNS.SendData(Data)
    If Mid(Data, Len(Data) - 1) = vbCrLf Then
        Debug.Print ("<<<: " & Mid(Data, 1, Len(Data) - 2))
    Else
        Debug.Print ("<<<: " & Data)
    End If
End Function

Public Function SetPersonalMessage(ByVal sMessage As String)
        Dim Message As String
        Dim EndResult As String
        Message = "<Data><PSM>" & sMessage & "</PSM><CurrentMedia></CurrentMedia></Data>"
        EndResult = "UUX # " & Len(Message) & vbCrLf & Message
        Call SendData(EndResult)
End Function

    'Set personal media for local user
Public Function SetCurrentMedia(ByVal sSong As String, ByVal sAlbum As String, ByVal sArtist As String)
        Dim Message As String
        Dim EndResult As String
        Message = "<Data><PSM></PSM><CurrentMedia>\0Music\01\0{0} - {1}\0" & sSong & "\0" & sArtist & "\0" & sAlbum & "\0\0</CurrentMedia></Data>"
        EndResult = "UUX # " & Len(Message) & vbCrLf & Message
        Call SendData(EndResult)
End Function

    'Clear all PSM data and music
Public Function ClearPSM_MUSIC()
        Dim Message As String
        Dim EndResult As String
        Message = "<Data><PSM></PSM><CurrentMedia></CurrentMedia></Data>"
        EndResult = "UUX # " & Len(Message) & vbCrLf & Message
        Call SendData(EndResult)
End Function
Public Function ChangeFriendlyname(ByVal sFriendlyname As String)
    Call SendData("PRP # " & "MFN " & Replace(sFriendlyname, " ", "%20") & vbCrLf)
End Function
