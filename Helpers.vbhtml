@Helper Copyright()
    ' Razor Portfolio VB.NET Version 1.0.0.0
    ' WWWeb Concepts Copyright 2017 (wwwebconcepts.com )
    ' James W. Threadgill james@wwwebconcepts.com
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
End Helper

@Helper SetTheme()
    ' Get template according to device or request
    Dim strTheme As String = HttpContext.Current.Application("Theme")
    Dim mTheme As String = HttpContext.Current.Application("MobileTheme")
    Dim SessionTheme As String = HttpContext.Current.Session.Item("SessionTheme")
    Dim ThemeCookie As String = ""
    If Not HttpContext.Current.Request.Cookies("ThemeCookie") Is Nothing Then
        ThemeCookie = HttpContext.Current.Request.Cookies("ThemeCookie").ToString
    End If

    If String.IsNullOrEmpty(SessionTheme) And String.IsNullOrEmpty(ThemeCookie) Then InitializeApplication.InitializeTheme()
    If ThemeCookie = "Mobile" Or SessionTheme = "Mobile" And Not String.IsNullOrEmpty(mTheme) Then
        strTheme = "~/" & mTheme
    Else
        strTheme = "~/" & strTheme
    End If
    @strTheme
End Helper

@Helper Canonical()
    ' Get page as desktop for use in desktop template
    Dim strProtocol As String = Request.Url.Scheme & "://"
    Dim qString As String = Request.ServerVariables("QUERY_STRING")
    Dim strCanonical As String = HttpContext.Current.Application("CanonicalDomain") & Strings.Replace(Request.ServerVariables("SCRIPT_NAME"), ".vbhtml", "")
    Try
        If Not String.IsNullOrEmpty(qString) Then
            If InStr(qString, "&Theme=d") Then
                qString = Strings.Replace(qString, "&Theme=d", "")
            ElseIf InStr(qString, "Theme=d") Then
                qString = Strings.Replace(qString, "Theme=d", "")
            ElseIf InStr(qString, "&Theme=m") Then
                qString = Strings.Replace(qString, "&Theme=m", "")
            ElseIf InStr(qString, "Theme=m") Then
                qString = Strings.Replace(qString, "Theme=m", "")
            End If
        End If

        If Not String.IsNullOrEmpty(qString) Then
            strCanonical = strProtocol & strCanonical & "?" & qString
        Else
            strCanonical = strProtocol & strCanonical
        End If
        strCanonical = URL.Rewrite(Server.HtmlEncode(strCanonical))
        @Html.Raw(strCanonical)
    Catch e As Exception
        'Dont show the error. 
    End Try
End Helper

@Helper GetBasePath()
    ' Get application base path 
    Dim strPath As String = ""
    Dim strProtocol As String = Request.Url.Scheme & "://"
    Dim port As String = Request.Url.Port
    If port <> 80 And port <> 433 Then port = ":" & port Else port = ""
    Dim strAppInstallFolder As String = InitializeApplication.AppInstallFolder()
    If Not String.IsNullOrEmpty(strAppInstallFolder) Then strAppInstallFolder &= "/"
    strPath = strProtocol & Request.ServerVariables("SERVER_NAME") & port & "/" & strAppInstallFolder
    @strPath
End Helper

@Helper PageAsMobile()             
    ' Get page link as mobile for use in desktop template
    Dim strPageM As String = Strings.Replace(Request.ServerVariables("SCRIPT_NAME"), ".vbhtml", "")
    Dim qString As String = Request.ServerVariables("QUERY_STRING")
    Dim mTheme As String = HttpContext.Current.Application("MobileTheme")
    Try
        If Not String.IsNullOrEmpty(mTheme) Then
            If Not String.IsNullOrEmpty(qString) Then
                strPageM = Strings.Replace(strPageM, ".vbhtml", "") & "?" & qString
                If InStr(qString, "Theme=d") Then
                    strPageM = Strings.Replace(strPageM, "Theme=d", "Theme=m")
                Else
                    strPageM &= "&Theme=m"
                End If
            Else
                strPageM &= "?Theme=m"
            End If
        Else
            strPageM = ""
        End If
        strPageM = URL.Rewrite(Server.HtmlEncode(strPageM))
        If InStr(strPageM, "?N=") Then strPageM = Replace(strPageM, "?Theme=", "&amp;Theme=")
        @Html.Raw(strPageM)
    Catch e As Exception
        'Dont show the error. 
    End Try
End Helper

@Helper PageAsDesktop()
    ' Get page link as desktop for use in mobile template
    Dim strPageD As String = Strings.Replace(Request.ServerVariables("SCRIPT_NAME"), ".vbhtml", "")
    Dim qString As String = Request.ServerVariables("QUERY_STRING")
    Try
        If Not String.IsNullOrEmpty(qString) Then
            strPageD = Strings.Replace(strPageD, ".vbhtml", "") & "?" & qString
            If InStr(qString, "Theme=m") Then
                strPageD = Strings.Replace(strPageD, "Theme=m", "Theme=d")
            Else
                strPageD &= "&Theme=d"
            End If
        Else
            strPageD &= "?Theme=d"
        End If
        strPageD = URL.Rewrite(Server.HtmlEncode(strPageD))
        If InStr(strPageD, "?N=") Then strPageD = Replace(strPageD, "?Theme=", "&amp;Theme=")
        @Html.Raw(strPageD)
    Catch e As Exception
        'Dont show the error. 
    End Try
End Helper