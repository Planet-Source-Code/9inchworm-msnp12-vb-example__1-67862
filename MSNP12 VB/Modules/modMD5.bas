Attribute VB_Name = "modMD5"
'-----------------------------------------------------------------
' baseMD5: Provide MD5 and SHA1 through cryptographic APIs
'-----------------------------------------------------------------

Option Explicit

'Functions
Private Declare Function CryptAcquireContext Lib "advapi32.dll" Alias "CryptAcquireContextA" (ByRef phProv As Long, ByVal pszContainer As String, ByVal pszProvider As String, ByVal dwProvType As Long, ByVal dwFlags As Long) As Long
Private Declare Function CryptReleaseContext Lib "advapi32.dll" (ByVal hProv As Long, ByVal dwFlags As Long) As Long
Private Declare Function CryptCreateHash Lib "advapi32.dll" (ByVal hProv As Long, ByVal Algid As Long, ByVal hKey As Long, ByVal dwFlags As Long, ByRef phHash As Long) As Long
Private Declare Function CryptDestroyHash Lib "advapi32.dll" (ByVal hHash As Long) As Long
Private Declare Function CryptHashData Lib "advapi32.dll" (ByVal hHash As Long, pbData As Any, ByVal dwDataLen As Long, ByVal dwFlags As Long) As Long
Private Declare Function CryptGetHashParam Lib "advapi32.dll" (ByVal hHash As Long, ByVal dwParam As Long, pbData As Any, pdwDataLen As Long, ByVal dwFlags As Long) As Long

'SHA1/MD5 consts
Private Const PROV_RSA_FULL = 1
Private Const ALG_CLASS_HASH = 32768
Private Const ALG_TYPE_ANY = 0
Private Const HP_HASHVAL = 2
Private Const HP_HASHSIZE = 4

'SHA1 consts
Private Const ALG_SID_SHA1 = 4
Private Const SHA1 = ALG_CLASS_HASH Or ALG_TYPE_ANY Or ALG_SID_SHA1

'MD5 consts
Private Const ALG_SID_MD5 = 3
Private Const CALG_MD5 = ((ALG_CLASS_HASH Or ALG_TYPE_ANY) Or ALG_SID_MD5)

Public Function MD5_Hex(ByVal Str As String) As String
    'Pass :)
    MD5_Hex = LCase$(CreateHash(Str, CALG_MD5))
End Function

Public Function SHA1_Hex(ByVal Str As String) As String
    'Pass :)
    SHA1_Hex = CreateHash(Str, SHA1)
End Function

Public Function CreateHash(ByVal Str As String, ByVal ConstVal As Long) As String
    'Create hash :)
    Dim hCtx As Long, hHash As Long, lRes As Long, lLen As Long, lIdx As Long, abData() As Byte

    'Get default provider context handle
    lRes = CryptAcquireContext(hCtx, vbNullString, vbNullString, PROV_RSA_FULL, 0)

    If lRes <> 0 Then
        'Create the hash
        lRes = CryptCreateHash(hCtx, ConstVal, 0, 0, hHash)
        If lRes <> 0 Then
            lRes = CryptHashData(hHash, ByVal Str, Len(Str), 0)
            If lRes <> 0 Then
                lRes = CryptGetHashParam(hHash, HP_HASHSIZE, lLen, 4, 0)
                If lRes <> 0 Then
                    ReDim abData(0 To lLen - 1)

                    'Get the hash value
                    lRes = CryptGetHashParam(hHash, HP_HASHVAL, abData(0), lLen, 0)
                    If lRes <> 0 Then
                        'Convert value to hex string
                        For lIdx = 0 To UBound(abData)
                            CreateHash = CreateHash & Right$("0" & Hex$(abData(lIdx)), 2)
                        Next
                    End If
                End If
            End If

            'Release the hash handle
            CryptDestroyHash hHash
        End If
    End If

    'Release the provider context
    CryptReleaseContext hCtx, 0
End Function

