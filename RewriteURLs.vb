'WWWeb Concepts Development Solutions wwwebconcepts.com
' Razor Portfolio ©2017 by James W. Threadgill  Created 5/22/2017
' james@wwwebconcepts.com with exceptions noted.
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
Imports System.Convert
Imports System.Web.HttpContext

Public Class URL

    ' Rewrite URLs. Get link case, use search and replace to rewrite the link 
    Public Shared Function Rewrite(ByVal URLString As String) As String
        Dim URLTempVal As String = URLString
        Dim URLRewrite As Boolean = InitializeApplication.RewriteURLS()

        ' Get URL Case 
        If (URLRewrite) Then
            Dim strURLCase As String = ""

            ' Get URL case
            If InStr(URLString, "PortfolioDetails") Then strURLCase = "Portfolio"

            ' Adjust query strings we're leaving unchanged. ' If we have a search query
            If InStr(URLString, "Query=") Then URLString = Replace(URLString, "&amp;Query=", "?Query=")
            ' If we have a a swap theme request but do not have a search query
            If InStr(URLString, "Theme=") And InStr(URLString, "Query=") = 0 Then URLString = Replace(URLString, "&amp;Theme=", "?Theme=")
            ' If there's pagination with swap theme
            If InStr(URLString, "Page=") And InStr(URLString, "Theme=") Then URLString = Replace(URLString, "?Theme=", "&amp;Theme=")

            ' Rewrite URLs to match rules in web.config
            Select Case strURLCase
                Case "Portfolio"
                    Rewrite = Replace(Replace(URLString, "?ItemID=", "/"), "&amp;Title=", "/")
                Case Else 'all other pages
                    Rewrite = Replace(URLString, "?Title=", "/")
            End Select
        Else
            Rewrite = (URLTempVal)
        End If

        Return Rewrite
    End Function

End Class
