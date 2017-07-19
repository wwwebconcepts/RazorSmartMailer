' WWWeb Concepts wwwebconcepts.com
' James W. Threadgill james@wwwebconcepts.com
' Copyright 2017
' Version 1.0.0.0
'=========================================================================================================================================================
' MIT License
' Copyright(c) 2017 James Threadgill
' Permission Is hereby granted, free Of charge, to any person obtaining a copy of this software And associated documentation files (the "Software"), to deal 
' in the Software without restriction, including without limitation the rights to use, copy, modify, merge, publish, distribute, sublicense, And/Or sell
' copies of the Software, And to permit persons to whom the Software Is'furnished to do so, subject to the following conditions:
'
' The above copyright notice And this permission notice shall be included In all copies Or substantial portions Of the Software.
'
' THE SOFTWARE Is PROVIDED "AS IS", WITHOUT WARRANTY Of ANY KIND, EXPRESS Or IMPLIED, INCLUDING BUT Not LIMITED To THE WARRANTIES Of MERCHANTABILITY, 
' FITNESS For A PARTICULAR PURPOSE And NONINFRINGEMENT. In NO Event SHALL THE  AUTHORS Or COPYRIGHT HOLDERS BE LIABLE For ANY CLAIM, DAMAGES Or OTHER
' LIABILITY, WHETHER In AN ACTION Of CONTRACT, TORT Or OTHERWISE, ARISING FROM, OUT OF Or IN CONNECTION WITH THE SOFTWARE Or THE USE Or OTHER DEALINGS 
' IN THE  SOFTWARE.
'=============================================================================================================================================================
Imports System.IO
Imports System.Net
Imports System.Text
Imports System.Convert
Imports System.Web.HttpContext
Imports System.Security.Cryptography
Imports System.Text.RegularExpressions.Regex

' site security utilities
Public Class ApplicationSecurity

    ' Filter user input for malicious code and remove
    Public Shared Function ValidateInput(ByVal inputString As String) As String

        If inputString <> "" Then
            inputString = Replace(inputString, "[^aA-zZ0-9]\s-\]", " ")
            inputString = Replace(inputString, "\s+", " ").Trim()
        End If

        Return inputString
    End Function

    ' base 64 querystring encoding for data security
    Const sBASE_64_CHARACTERS = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+-"

    'Encode  
    Public Shared Function Encode(ByVal asContents As String) As String

        If InitializeApplication.EncodeIDs() Then
            Dim lnPosition As Integer
            Dim lsResult As String
            Dim Char1 As String
            Dim Char2 As String
            Dim Char3 As String
            Dim Char4 As String
            Dim Byte1 As String
            Dim Byte2 As String
            Dim Byte3 As String
            Dim SaveBits1 As String
            Dim SaveBits2 As String
            Dim lsGroupBinary As String
            Dim lsGroup64 As String

            If Len(asContents) Mod 3 > 0 Then asContents = asContents & StrDup(3 - (Len(asContents) Mod 3), " ")
            lsResult = ""

            For lnPosition = 1 To Len(asContents) Step 3
                lsGroup64 = ""
                lsGroupBinary = Mid(asContents, lnPosition, 3)

                Byte1 = Asc(Mid(lsGroupBinary, 1, 1)) : SaveBits1 = Byte1 And 3
                Byte2 = Asc(Mid(lsGroupBinary, 2, 1)) : SaveBits2 = Byte2 And 15
                Byte3 = Asc(Mid(lsGroupBinary, 3, 1))

                Char1 = Mid(sBASE_64_CHARACTERS, ((Byte1 And 252) \ 4) + 1, 1)
                Char2 = Mid(sBASE_64_CHARACTERS, (((Byte2 And 240) \ 16) Or (SaveBits1 * 16) And &HFF) + 1, 1)
                Char3 = Mid(sBASE_64_CHARACTERS, (((Byte3 And 192) \ 64) Or (SaveBits2 * 4) And &HFF) + 1, 1)
                Char4 = Mid(sBASE_64_CHARACTERS, (Byte3 And 63) + 1, 1)
                lsGroup64 = Char1 & Char2 & Char3 & Char4

                lsResult &= lsGroup64
            Next
            Return lsResult
        Else
            Return asContents
        End If
    End Function

    'decode
    Public Shared Function Decode(ByVal asContents As String) As String

        If InitializeApplication.EncodeIDs() Then
            Dim lnPosition As Integer
            Dim lsResult As String
            Dim Char1 As String
            Dim Char2 As String
            Dim Char3 As String
            Dim Char4 As String
            Dim Byte1 As String
            Dim Byte2 As String
            Dim Byte3 As String
            Dim lsGroupBinary As String
            Dim lsGroup64 As String

            If Len(asContents) Mod 4 > 0 Then asContents = asContents & StrDup(4 - (Len(asContents) Mod 4), " ")
            lsResult = ""

            For lnPosition = 1 To Len(asContents) Step 4
                lsGroupBinary = ""
                lsGroup64 = Mid(asContents, lnPosition, 4)
                Char1 = InStr(sBASE_64_CHARACTERS, Mid(lsGroup64, 1, 1)) - 1
                Char2 = InStr(sBASE_64_CHARACTERS, Mid(lsGroup64, 2, 1)) - 1
                Char3 = InStr(sBASE_64_CHARACTERS, Mid(lsGroup64, 3, 1)) - 1
                Char4 = InStr(sBASE_64_CHARACTERS, Mid(lsGroup64, 4, 1)) - 1
                Byte1 = Chr(((Char2 And 48) \ 16) Or (Char1 * 4) And &HFF)
                Byte2 = lsGroupBinary & Chr(((Char3 And 60) \ 4) Or (Char2 * 16) And &HFF)
                Byte3 = Chr((((Char3 And 3) * 64) And &HFF) Or (Char4 And 63))
                lsGroupBinary = Byte1 & Byte2 & Byte3

                lsResult &= lsGroupBinary
            Next
            Return lsResult
        Else
            Return asContents
        End If
    End Function
End Class
