' ReCAPTCHANet VB.NET Version 1.0.0.0
' (Based on DMXzone reCAPTCHA VB Script DMXzone.com)
' James W. Threadgill james@wwwebconcepts.com
' Copyright 2017 WWWeb Concepts wwwebconcepts.com
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
' calling code
'Dim theCAPTCHA As New ReCAPTCHANet
'theCAPTCHA.FormName = "form1"
'theCAPTCHA.PrivateKey = "" & ReCAPTCHAPrivateKey & ""
'theCAPTCHA.PublicKey = "" & ReCAPTCHAPublicKey & ""
'theCAPTCHA.QueryParameter = ""
'theCAPTCHA.RedirectURL = ""
'theCAPTCHA.FailURL = ""
'theCAPTCHA.Construct() As boolean
'<div class="recaptcha"><%=theCAPTCHA.getControl("" & ReCAPTCHATheme & "","" & ReCAPTCHALanguage & "")%></div>
'--------------------------------------------------------
Imports System.IO
Imports System.Net
Imports System.Web.HttpContext
Public Class ReCAPTCHANet
    Public ReadOnly Property ErrorCodes As New List(Of WebException)

    Private m_PrivateKey As String
    Private m_PublicKey As String
    Private m_RedirectURL As String
    Private m_FailURL As String
    Private m_QueryParameter As String
    Private m_recaptcha_challenge_field As String
    Private m_recaptcha_response_field As String
    Private m_error As String

    Public Property PrivateKey() As String
        Get
            Return m_PrivateKey
        End Get
        Set(ByVal value As String)
            m_PrivateKey = value
        End Set
    End Property
    Public Property PublicKey() As String
        Get
            Return m_PublicKey
        End Get
        Set(ByVal value As String)
            m_PublicKey = value
        End Set
    End Property
    Public Property RedirectURL() As String
        Get
            Return m_RedirectURL
        End Get
        Set(ByVal value As String)
            m_RedirectURL = value
        End Set
    End Property
    Public Property FailURL() As String
        Get
            Return m_FailURL
        End Get
        Set(ByVal value As String)
            m_FailURL = value
        End Set
    End Property
    Public Property QueryParameter() As String
        Get
            Return m_QueryParameter
        End Get
        Set(ByVal value As String)
            m_QueryParameter = value
        End Set
    End Property

    ' Initialize Class
    Public Sub New()
        m_PrivateKey = ""
        m_PublicKey = ""
        m_RedirectURL = ""
        m_FailURL = ""
        m_QueryParameter = ""
        m_error = ""
    End Sub

    Public Shared Function GetUserIPAddress() As String
        Dim ip As String = Current.Request.ServerVariables("HTTP_X_FORWARDED_FOR")
        If Not String.IsNullOrEmpty(ip) Then
            Dim ipRange As String() = ip.Split(","c)
            ip = ipRange(0)
        Else
            ip = Current.Request.ServerVariables("REMOTE_ADDR")
        End If
        Return ip.Trim()
    End Function

    Public Function Construct() As Boolean
        Dim server_Response As Boolean = False
        Dim redirect As String = m_RedirectURL

        If Not String.IsNullOrEmpty(Current.Request("g-recaptcha-response")) Then
            ' robots version
            m_recaptcha_response_field = Current.Request("g-recaptcha-response")
            server_Response = ValidateReCaptcha()
        ElseIf Not String.IsNullOrEmpty(Current.Request("recaptcha_response_field")) Then
            ' classic version
            m_recaptcha_challenge_field = Current.Request("recaptcha_challenge_field")
            m_recaptcha_response_field = Current.Request("recaptcha_response_field")
            server_Response = Recaptcha_Confirm()
        End If

        'If we have a response and redirect enabled.
        If Not String.IsNullOrEmpty(redirect) And Not String.IsNullOrEmpty(m_recaptcha_response_field) Then
            m_error = ""
            Dim qStringOperator As String = "?"
            If String.IsNullOrEmpty(m_QueryParameter) Then m_QueryParameter = "Recaptcha"
            If InStr(redirect, "?") Then qStringOperator = ("&")
            If Not (server_Response) And String.IsNullOrEmpty(m_FailURL) Then
                ' incorrect
                m_error = "incorrect-captcha-sol"
                Current.Response.Redirect(redirect & qStringOperator & m_QueryParameter & "=" & m_error)
            ElseIf Not (server_Response) And Not String.IsNullOrEmpty(m_FailURL) Then
                ' incorrect with optional fail URL
                m_error = "incorrect-captcha-sol"
                Current.Response.Redirect(m_FailURL & qStringOperator & m_QueryParameter & "=" & m_error)
            Else
                ' correct
                Current.Response.Redirect(redirect & qStringOperator & m_QueryParameter & "=true")
            End If
        End If

        Return (server_Response)
    End Function

    Public Overloads Function GetControl() As String
        GetControl = "<script src=""https://www.google.com/recaptcha/api.js"" async defer></script>" & vbCrLf &
                     "<div class=""g-recaptcha"" data-sitekey=""" & m_PublicKey & """></div>" & vbCrLf
        Return GetControl
    End Function

    Public Overloads Function GetControl(ByVal theme As String, ByVal language As String) As String
        Dim errString As String
        errString = ""

        If String.IsNullOrEmpty(theme) Then theme = "red"
        If String.IsNullOrEmpty(language) Then language = "en"
        If Not String.IsNullOrEmpty(m_error) Then errString = "&amp;error=" & m_error

        GetControl =
      "<script type=""text/javascript"">" &
      "var RecaptchaOptions = {" &
      "   theme : '" & theme & "'," &
        "   lang : '" & language & "'," &
      "   tabindex : 0" &
      "};" &
      "</script>" & vbCrLf &
      "<script type=""text/javascript"" src=""http://www.google.com/recaptcha/api/challenge?k=" & m_PublicKey & errString & """></script>" & vbCrLf &
      "<noscript>" & vbCrLf &
        "<iframe src=""http://www.google.com/recaptcha/api/noscript?k=" & m_PublicKey & """ frameborder=""1""></iframe>" & vbCrLf &
          "<textarea name=""recaptcha_challenge_field"" rows=""3"" cols=""40""></textarea>" & vbCrLf &
          "<input type=""hidden"" name=""recaptcha_response_field"" value=""manual_challenge"" />" & vbCrLf &
      "</noscript>" & vbCrLf

        Return GetControl
    End Function

    ' returns True if correct
    Private Function Recaptcha_Confirm() As Boolean

        Dim VarString As String
        VarString =
              "?privatekey=" & m_PrivateKey &
              "&remoteip=" & GetUserIPAddress() &
              "&challenge=" & m_recaptcha_challenge_field &
              "&response=" & Current.Server.UrlEncode(m_recaptcha_response_field)
        Try
            Dim objXMLHttp As WebRequest = WebRequest.Create("http://www.google.com/recaptcha/api/verify" & VarString)
            objXMLHttp.Method = "POST"
            objXMLHttp.ContentType = "application/x-www-form-urlencoded"
            objXMLHttp.Timeout = 5 * 1000 ' 5 Seconds to avoid getting locked up

            Dim writer As StreamWriter = New StreamWriter(objXMLHttp.GetRequestStream())
            writer.Write(objXMLHttp)
            writer.Close()

            Dim response As WebResponse = objXMLHttp.GetResponse()
            Dim reader As StreamReader = New StreamReader(response.GetResponseStream())
            Dim reCAPTCHAResponse As String = reader.ReadLine()
            reader.Close()
            response.Close()

            If LCase(reCAPTCHAResponse) = "true" Then
                Return True
            End If
        Catch e As WebException
            ErrorCodes.Add(e)
        End Try

        Return False
    End Function

    ' Returns True if correct
    Private Function ValidateReCaptcha() As Boolean

        Dim VarString As String
        VarString =
              "?secret=" & m_PrivateKey &
              "&response=" & Current.Server.UrlEncode(m_recaptcha_response_field)
        Try
            If Not String.IsNullOrEmpty(m_recaptcha_response_field) Then
                Dim objJsonHttp As WebRequest = WebRequest.Create("https://www.google.com/recaptcha/api/siteverify" & VarString)
                objJsonHttp.Method = "POST"
                objJsonHttp.ContentType = "application/json; charset=utf-8"
                objJsonHttp.Timeout = 5 * 1000 ' 5 Seconds to avoid getting locked up

                Dim writer As StreamWriter = New StreamWriter(objJsonHttp.GetRequestStream())
                writer.Write(objJsonHttp)
                writer.Close()

                Dim response As WebResponse = objJsonHttp.GetResponse()
                Dim reader As New StreamReader(response.GetResponseStream())
                Dim reCAPTCHAResponse As String = reader.ReadToEnd()
                reader.Close()
                response.Close()

                If InStr(reCAPTCHAResponse, """success"": true") Then
                    Return True
                End If
            End If
        Catch e As WebException
            ErrorCodes.Add(e)
        End Try

        Return False
    End Function

End Class




