' WWWeb Concepts wwwebconcepts.com
' James W. Threadgill james@wwwebconcepts.com
' Copyright 2002-2017
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
Imports System.Math
Imports System.Convert
Imports WebMatrix.Data
Imports System.DateTime
Imports System.Web.HttpContext
Imports System.Text.RegularExpressions.Regex
Imports System.Configuration.ConfigurationManager

Public Class InitializeApplication

    ' encapsulate app settings from web.config as properties
    Public LoggedInUserID As Long
    Public LoggedInUserName As String
    Public LoggedInGroupID As Long
    Public SecurityProfile() As Boolean

    Public Shared ReadOnly Property DataConn() As String
        Get
            Return AppSettings("userDatabase")
        End Get
    End Property
    Public Shared ReadOnly Property SmtpServer() As String
        Get
            Return AppSettings("SmtpServer")
        End Get
    End Property
    Public Shared ReadOnly Property EnableSsl() As String
        Get
            Return AppSettings("EnableSsl")
        End Get
    End Property
    Public Shared ReadOnly Property UserName() As String
        Get
            Return AppSettings("UserName")
        End Get
    End Property
    Public Shared ReadOnly Property Password() As String
        Get
            Return AppSettings("Password")
        End Get
    End Property
    Public Shared ReadOnly Property From() As String
        Get
            Return AppSettings("From")
        End Get
    End Property

    Public Shared ReadOnly Property eMailContentEncoding() As String
        Get
            Return AppSettings("eMailContentEncoding")
        End Get
    End Property

    Public Shared ReadOnly Property RecaptchaPublicKey() As String
        Get
            Return AppSettings("recaptchaPublicKey")
        End Get
    End Property
    Public Shared ReadOnly Property RecaptchaPrivateKey() As String
        Get
            Return AppSettings("recaptchaPrivateKey")
        End Get
    End Property
    Public Shared ReadOnly Property AppInstallFolder() As String
        Get
            Return AppSettings("AppInstallFolder")
        End Get
    End Property
    Public Shared ReadOnly Property RewriteURLS() As String
        Get
            Return AppSettings("RewriteURLS")
        End Get
    End Property
    Public Shared ReadOnly Property EncodeIDs() As Boolean
        Get
            Return AppSettings("EncodeIDs")
        End Get
    End Property

    Public Shared Sub InitializeAppVars()
        Dim appInit As Boolean = True
        ' TODO: load Application variables and values fron database

        ' forms
        Current.Application("ReCAPTCHATheme") = "blackglass"
        Current.Application("ReCAPTCHALanguage") = "en"
        Current.Application("ContactFormSubject") = "Contact Us"
        Current.Application("RecipientEmail") = "james@wwwebconcepts.com"
        Current.Application("RecipientEmailName") = "Razor Smart Mailer"
        Current.Application("FormInstructionsLbl") = "All fields required."
        Current.Application("FormNameLbl") = "Name"
        Current.Application("FormAddressLbl") = "Address"
        Current.Application("FormCityLbl") = "City"
        Current.Application("FormStateLbl") = "State"
        Current.Application("FormZipCodeLbl") = "Zip"
        Current.Application("FormPhoneLbl") = "Phone"
        Current.Application("FormEmailLbl") = "eMail Address"
        Current.Application("FormMessageLbl") = "Message"
        Current.Application("FormBtnTxt") = "Contact Us"
        Current.Application("ReCaptchaLbl") = "CAPTCHA Verification"
        Current.Application("FormFileLbl") = "Attachments"
        Current.Application("eMailAttachments") = "Files Attached: "

        ' application
        Current.Application("RewriteURLS") = True
        Current.Application("Theme") = "_SiteLayout.vbhtml"
        Current.Application("MobileTheme") = "_MobileSiteLayout.vbhtml"
        Current.Application("EmailTheme") = "_EmailLayout.vbhtml"

        Current.Application("CanonicalDomain") = "localhost:61341"

        ' seo
        Current.Application("DefaultTitle") = ("The default site title from application variable")
        Current.Application("DefaultKeywords") = "the the default string of keywords"
        Current.Application("DefaultDescription") = "the the default string description"

        ' pagination
        Current.Application("FeaturedNumRows") = 1
        Current.Application("FeaturedNumCols") = 2
        Current.Application("SearchNumRows") = 1
        Current.Application("SearchNumCols") = 2

        ' rs nav bar strings
        Current.Application("RSNavStyle") = "blue"
        Current.Application("RSFirstTxt") = "First"
        Current.Application("RSPrevTxt") = "Previous"
        Current.Application("RSNextTxt") = "Next"
        Current.Application("RSLastTxt") = "Last"
        Current.Application("RSNavCustomFirstImg") = ""
        Current.Application("RSNavCustomPreviousImg") = ""
        Current.Application("RSNavCustomNextImg") = ""
        Current.Application("RSNavCustomLastImg") = ""

        ' social
        Current.Application("AddThisToolBar") = ""

        ' headers
        Current.Application("SearchTitle") = "Search"
        Current.Application("ContactFormHdr") = "Contact Us"
        Current.Application("ContactFormSubHdr") = "Let us solve your problem."
        Current.Application("ContactThanksHdr") = "Thanks for getting in touch."
        Current.Application("ContactThanksSubHdr") = "You submitted the following:"
        Current.Application("ContactFormSubject") = "Website Contact Form Submission"

        ' trimmed text lengths (int)
        Current.Application("FeaturedTextTrim") = 175
        Current.Application("SearchTextTrim") = 175

        ' format
        Current.Application("DateFormat") = "MMMM dd yyyy"

        ' set app initialized
        Current.Application("AppInitialized") = appInit

    End Sub

    Public Shared Sub InitializeAdminAppVars()
        Dim adminInit As Boolean = True
        Current.Application("AdminInitialized") = adminInit
    End Sub

    Public Shared Property ThemeCookie As HttpCookie = New HttpCookie("ThemeCookie")

    Public Shared Sub InitializeTheme()
        If Current.Request.Cookies("ThemeCookie") Is Nothing Or Current.Session.Item("SessionTheme") Is Nothing Then
            ' dimensions variables, detect mobile device 
            Dim strTheme As String = "Default"
            Dim Is_Mobile As Boolean = DetectMobileDevice()
            If Is_Mobile And Not String.IsNullOrEmpty(Current.Application("MobileTheme")) Then
                strTheme = "Mobile"
            End If
            ThemeCookie.Value = strTheme
            ThemeCookie.Expires = Now.AddDays(30)
            Current.Response.Cookies.Add(ThemeCookie)
            Current.Session.Add("SessionTheme", strTheme)
        End If
    End Sub

    Public Shared Sub SwapTheme(ByVal Theme As String)
        ' reset theme cookie/session values 
        Dim strTheme As String = "Default"
        If Theme = "m" And Not String.IsNullOrEmpty(Current.Application("MobileTheme")) Then
            strTheme = "Mobile"
        End If
        ThemeCookie.Value = strTheme
        ThemeCookie.Expires = Now.AddDays(30)
        Current.Response.Cookies.Add(ThemeCookie)
        Current.Session.Item("SessionTheme") = strTheme
    End Sub

    'detect mobile device and returns true or null. exits when mobile device discovered.
    Public Shared Function DetectMobileDevice() As Boolean
        Dim mobile_agents, user_agent, mobile_ua As String

        ' Check server variables
        If InStr(Current.Request.ServerVariables("HTTP_ACCEPT"), "application/vnd.wap.xhtml+xml") Then
            Return True
        ElseIf Not Current.Request.ServerVariables("HTTP_X_PROFILE") Is Nothing Then
            Return True
        ElseIf Not Current.Request.ServerVariables("HTTP_PROFILE") Is Nothing Then
            Return True
        End If

        ' If we haven't detected a mobile device request user agent.
        If Not Current.Request.UserAgent Is Nothing Then
            user_agent = Current.Request.UserAgent
        Else
            user_agent = Current.Request.ServerVariables("HTTP_USER_AGENT")
        End If

        ' Regex search for matching user_agent.
        If Not String.IsNullOrEmpty(user_agent) Then
            Dim pattern As String = "(up.browser|up.link|mmp|symbian|smartphone|midp|wap|phone|windows ce|pda|mobile|mini|palm|ipad)"
            Dim options As RegexOptions = RegexOptions.IgnoreCase
            Dim match As Match = Regex.Match(user_agent, pattern, options)
            If match.Length > 0 Then
                Return True
            End If

            ' Text match against array of known mobile_agents
            mobile_agents = ("acs- alav alca amoi audi avan benq bird blac blaz brew cell cldc cmd- dang doco eric hipt inno ipaq java jigs kddi keji leno lg-c lg-d lg-g lge- maui maxo midp mits mmef mobi mot- moto mwbp nec- newt noki oper palm pana pant phil play port prox qwap sage sams sany sch- sec- send seri sgh- shar sie- siem smal smar sony sph- symb t-mo teli tim- tosh tsm- upg1 upsi vk-v voda wap- wapa wapi wapp wapr webc winw winw xda xda- ")
            Dim Devices As Array = Split(mobile_agents, " ")
            mobile_agents = ""
            mobile_ua = LCase(Left(user_agent, 4))
            For i = 0 To UBound(Devices) - 1
                If mobile_ua = Devices(i) Then
                    Return True
                End If
            Next
        End If
        Return False
    End Function
End Class

