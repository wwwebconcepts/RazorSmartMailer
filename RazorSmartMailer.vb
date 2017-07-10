'=========================================================
' WWWeb Concepts wwwebconcepts.com
' James W. Threadgill james@wwwebconcepts.com
' RazorSmartMailer Version 1.0.0.0 Copyright 2017
'=========================================================

' RazorSmartMailer calling code
' Dim theMailer As New RazorSmartMailer
' RazorSmartMailer templater properties
'theMailer.MailTemplatePath = "MailTemplatePath"
'theMailer.PreMailerCss = true
'theMailer.AddHTMLBasePath = false
'theMailer.ParsePaths = false
'theMailer.ErrorCodes = nothing (returns errors)
'theMailer.GetMailBody() ' returns body only

' RazorSmartMailer sendmail properties
'theMailer.EmailRecipient = ""
'theMailer.EmailFrom = ""
'theMailer.EmailCC = ""
'theMailer.EmailBC = ""
'theMailer.EmailReplyTo = ""
'theMailer.EMailEncoding = "utf-8"
'theMailer.EMailSubject = ""
'theMailer.IsBodyHtml = True
'theMailer.EmailPriority  = "Normal"
'theMailer.SystemMailHeaders = ""
'theMailer.AdditionalHeaders = ""
'theMailer.AttachmentFolder = "SmartMailerAttachments"
'theMailer.SaveAttachments = True
'theMailer.SendWebMail() ' does it all with web mail helper

' SMTP Server properites
'theMailer.SmtpUsername = ""
'theMailer.SmtpPassword = ""
'theMailer.SmtpHost = "localhost"
'theMailer.SmtpEnableSsl = False
'theMailer.SmtpPort = 25
'theMailer.SystemMailHeaders = Nothing
'theMailer.SystemMailEncoding = Encding.UTF8
'theMailer.SendSystemMail() ' does it all with system mail
' Resize
'theMailer.ImageSizes = (String Format "width, height, suffix | width, height, suffix |)
'theMailer.PreventEnlarge =Ttrue
'theMailer.PreserveAspectRatio = True
' Crop
'theMailer.CropSizes = (String Format: "width, height | width, height |) *Must match resize order.
'theMailer.CropPosition = (default "center-middle")
' Watermark
'theMailer.WatermarkMask = ""
'theMailer.WatermarkPadding = 10
'theMailer.WatermarkOpacity = 50
'theMailer.WatermarkSizes = "128, 128"
'theMailer.WatermarkAlign = "Center-Middle"
' Captions
'theMailer.CaptionText = ""
'theMailer.CaptionFont = "Ariel"
'theMailer.CaptionFontSizes = "14"
'theMailer.CaptionFontColor = "Black"
'theMailer.CaptionFontStyle =  "Bold" 'Valid values are: "Regular", "Bold", "Italic", "Underline", and "Strikeout".
'theMailer.CaptionOpacity = 100
'theMailer.CaptionPadding = 10
'theMailer.CaptionAlign = "Center-Middle" 
' Uploader
'theMailer.UploadFolder = "SmartMailerUploads" 

Imports PreMailer
Imports System.IO
Imports System.Net
Imports System.Math
Imports PreMailer.Net
Imports System.IO.Path
Imports System.Net.Mail
Imports System.Web.Helpers
Imports System.Web.HttpContext
Imports System.Web.Helpers.WebMail
Imports Microsoft.VisualBasic.Strings
Public Class RazorSmartMailer
    Public ReadOnly Property ErrorCodes As New List(Of WebException)

    ' Section templater
    '***********************************************************************************
    ' Templater Private variables
    '***********************************************************************************
    Private p_basePath As String
    Private p_appInstallFolder As String
    Private p_mailTemplatePath As String
    Private p_addHTMLBasePath As Boolean
    Private p_parsePaths As Boolean
    Private p_preMailerCss As Boolean
    ' Templater properties
    Public Property AppInstallFolder() As String
        Get
            Return p_appInstallFolder
        End Get
        Set(ByVal value As String)
            p_appInstallFolder = value
        End Set
    End Property
    Public Property MailTemplatePath() As String
        Get
            Return p_mailTemplatePath
        End Get
        Set(ByVal value As String)
            p_mailTemplatePath = value
        End Set
    End Property
    Public Property AddHTMLBasePath() As Boolean
        Get
            Return p_addHTMLBasePath
        End Get
        Set(ByVal value As Boolean)
            p_addHTMLBasePath = value
        End Set
    End Property
    Public Property ParsePaths() As Boolean
        Get
            Return p_parsePaths
        End Get
        Set(ByVal value As Boolean)
            p_parsePaths = value
        End Set
    End Property
    Public Property PreMailerCss() As Boolean
        Get
            Return p_preMailerCss
        End Get
        Set(ByVal value As Boolean)
            p_preMailerCss = value
        End Set
    End Property

    ' Section eMail
    '***********************************************************************************
    ' eMail Private variables
    '***********************************************************************************
    Private p_emailRecipient As String
    Private p_emailFrom As String
    Private p_emailCC As String
    Private p_emailBCC As String
    Private p_emailReplyTo As String
    Private p_eMailEncoding As String
    Private p_systemMailEncoding As Encoding
    Private p_eMailSubject As String
    Private p_isBodyHtml As Boolean
    Private p_emailPriority As String
    Private p_additionalHeaders As New List(Of String)
    Private p_systemMailHeaders As New NameValueCollection()
    Private p_filesToAttach As HttpFileCollection
    Private p_imagesUploaded As New List(Of String)
    Private p_attachmentFolder As String
    Private p_attachmentMaxLength As Integer
    Private p_saveAttachments As Boolean
    Private p_fileNamingStyle As String
    Private p_successRedirect As String
    ' SMTP
    Private p_SmtpUserName As String
    Private p_SmtpPassword As String
    Private p_SmtpHost As String
    Private p_SmtpEnableSsl As Boolean
    Private p_SmtpPort As Integer
    ' Shared variable p_emailAttachments contains attachment file paths.
    Private Shared p_emailAttachments As New List(Of String)

    ' eMail properties
    Public Property SmtpUserName() As String
        Get
            Return p_SmtpUserName
        End Get
        Set(ByVal value As String)
            p_SmtpUserName = value
        End Set
    End Property
    Public Property SmtpPassword() As String
        Get
            Return p_SmtpPassword
        End Get
        Set(ByVal value As String)
            p_SmtpPassword = value
        End Set
    End Property
    Public Property SmtpHost() As String
        Get
            Return p_SmtpHost
        End Get
        Set(ByVal value As String)
            p_SmtpHost = value
        End Set
    End Property
    Public Property SmtpEnableSsl() As String
        Get
            Return p_SmtpEnableSsl
        End Get
        Set(ByVal value As String)
            p_SmtpEnableSsl = value
        End Set
    End Property
    Public Property SmtpPort() As String
        Get
            Return p_SmtpPort
        End Get
        Set(ByVal value As String)
            p_SmtpPort = value
        End Set
    End Property
    Public Property eMailRecipient() As String
        Get
            Return p_emailRecipient
        End Get
        Set(ByVal value As String)
            p_emailRecipient = value
        End Set
    End Property
    Public Property eMailFrom() As String
        Get
            Return p_emailFrom
        End Get
        Set(ByVal value As String)
            p_emailFrom = value
        End Set
    End Property
    Public Property eMailCC() As String
        Get
            Return p_emailCC
        End Get
        Set(ByVal value As String)
            p_emailCC = value
        End Set
    End Property
    Public Property eMailBCC() As String
        Get
            Return p_emailBCC
        End Get
        Set(ByVal value As String)
            p_emailBCC = value
        End Set
    End Property
    Public Property eMailReplyTo() As String
        Get
            Return p_emailReplyTo
        End Get
        Set(ByVal value As String)
            p_emailReplyTo = value
        End Set
    End Property
    Public Property eMailEncoding() As String
        Get
            Return p_eMailEncoding
        End Get
        Set(ByVal value As String)
            p_eMailEncoding = value
        End Set
    End Property
    Public Property SystemMailEncoding() As Encoding
        Get
            Return p_systemMailEncoding
        End Get
        Set(ByVal value As Encoding)
            p_systemMailEncoding = value
        End Set
    End Property
    Public Property eMailSubject() As String
        Get
            Return p_eMailSubject
        End Get
        Set(ByVal value As String)
            p_eMailSubject = value
        End Set
    End Property
    Public Property IsBodyHtml() As Boolean
        Get
            Return p_isBodyHtml
        End Get
        Set(ByVal value As Boolean)
            p_isBodyHtml = value
        End Set
    End Property
    Public Property eMailPriority() As String
        Get
            Return p_emailPriority
        End Get
        Set(ByVal value As String)
            p_emailPriority = value
        End Set
    End Property
    Public Property AdditionalHeaders() As List(Of String)
        Get
            Return p_additionalHeaders
        End Get
        Set(ByVal value As List(Of String))
            p_additionalHeaders = value
        End Set
    End Property
    Public Property SystemMailHeaders() As NameValueCollection
        Get
            Return p_systemMailHeaders
        End Get
        Set(ByVal value As NameValueCollection)
            p_systemMailHeaders = value
        End Set
    End Property

    Public Property AttachmentFolder() As String
        Get
            Return p_attachmentFolder
        End Get
        Set(ByVal value As String)
            p_attachmentFolder = value
        End Set
    End Property

    Public Property AttachmentMaxLength() As Integer
        Get
            Return p_attachmentMaxLength
        End Get
        Set(ByVal value As Integer)
            p_attachmentMaxLength = value
        End Set
    End Property
    Public Property FileNamingStyle() As String
        Get
            Return p_fileNamingStyle
        End Get
        Set(ByVal value As String)
            p_fileNamingStyle = value
        End Set
    End Property
    Public Property SaveAttachments() As Boolean
        Get
            Return p_saveAttachments
        End Get
        Set(ByVal value As Boolean)
            p_saveAttachments = value
        End Set
    End Property
    Public Property SuccessRedirect() As String
        Get
            Return p_successRedirect
        End Get
        Set(ByVal value As String)
            p_successRedirect = value
        End Set
    End Property
    Public Shared ReadOnly Property eMailAttachments() As List(Of String)
        Get
            Return p_emailAttachments
        End Get
    End Property

    ' Section Imaging
    '***********************************************************************************
    ' Imaging Private variables
    '***********************************************************************************
    ' Resize
    Private p_imageSizes As String
    Private p_preventEnlarge As Boolean
    Private p_preserveAspectRatio As Boolean
    ' Crop
    Private p_cropSizes As String
    Private p_cropPosition As String
    ' Watermark
    Private p_watermarkPadding As Integer
    Private p_watermarkMask As String
    Private p_watermarkOpacity As Integer
    Private p_watermarkAlign As String
    Private p_watermarkSizes As String
    ' Caption
    Private p_captionText As String
    Private p_captionFont As String
    Private p_captionFontSizes As String
    Private p_captionFontColor As String
    Private p_captionFontStyle As String
    Private p_captionOpacity As Integer
    Private p_captionAlign As String
    Private p_captionPadding As Integer

    ' Imaging Public Properties Resize
    Public Property ImageSizes() As String
        Get
            Return p_imageSizes
        End Get
        Set(ByVal value As String)
            p_imageSizes = value
        End Set
    End Property
    Public Property PreventEnlarge() As Boolean
        Get
            Return p_preventEnlarge
        End Get
        Set(ByVal value As Boolean)
            p_preventEnlarge = value
        End Set
    End Property
    Public Property PreserveAspectRatio() As Boolean
        Get
            Return p_preserveAspectRatio
        End Get
        Set(ByVal value As Boolean)
            p_preserveAspectRatio = value
        End Set
    End Property
    ' Crop
    Public Property CropSizes() As String
        Get
            Return p_cropSizes
        End Get
        Set(ByVal value As String)
            p_cropSizes = value
        End Set
    End Property
    Public Property CropPosition() As String
        Get
            Return p_cropPosition
        End Get
        Set(ByVal value As String)
            p_cropPosition = value
        End Set
    End Property
    ' Watermark
    Public Property WatermarkMask() As String
        Get
            Return p_watermarkMask
        End Get
        Set(ByVal value As String)
            p_watermarkMask = value
        End Set
    End Property
    Public Property WatermarkAlign() As String
        Get
            Return p_watermarkAlign
        End Get
        Set(ByVal value As String)
            p_watermarkAlign = value
        End Set
    End Property
    Public Property WatermarkSizes() As String
        Get
            Return p_watermarkSizes
        End Get
        Set(ByVal value As String)
            p_watermarkSizes = value
        End Set
    End Property
    Public Property WatermarkPadding() As Integer
        Get
            Return p_watermarkPadding
        End Get
        Set(ByVal value As Integer)
            p_watermarkPadding = value
        End Set
    End Property
    Public Property WatermarkOpacity() As Integer
        Get
            Return p_watermarkOpacity
        End Get
        Set(ByVal value As Integer)
            p_watermarkOpacity = value
        End Set
    End Property
    ' Caption
    Public Property CaptionText() As String
        Get
            Return p_captionText
        End Get
        Set(ByVal value As String)
            p_captionText = value
        End Set
    End Property
    Public Property CaptionFont() As String
        Get
            Return p_captionFont
        End Get
        Set(ByVal value As String)
            p_captionFont = value
        End Set
    End Property
    Public Property CaptionFontSizes() As String
        Get
            Return p_captionFontSizes
        End Get
        Set(ByVal value As String)
            p_captionFontSizes = value
        End Set
    End Property
    Public Property CaptionFontColor() As String
        Get
            Return p_captionFontColor
        End Get
        Set(ByVal value As String)
            p_captionFontColor = value
        End Set
    End Property
    Public Property CaptionFontStyle() As String
        Get
            Return p_captionFontStyle
        End Get
        Set(ByVal value As String)
            p_captionFontStyle = value
        End Set
    End Property
    Public Property CaptionOpacity() As Integer
        Get
            Return p_captionOpacity
        End Get
        Set(ByVal value As Integer)
            p_captionOpacity = value
        End Set
    End Property
    Public Property CaptionAlign() As String
        Get
            Return p_captionAlign
        End Get
        Set(ByVal value As String)
            p_captionAlign = value
        End Set
    End Property
    Public Property CaptionPadding() As Integer
        Get
            Return p_captionPadding
        End Get
        Set(ByVal value As Integer)
            p_captionPadding = value
        End Set
    End Property

    '**********************************
    ' File uploader/image processor
    '**********************************
    Private p_uploadFolder As String
    ' List of images processed by imaging
    Private Shared p_imageArray As New List(Of String)
    ' List of all uploaded files.
    Private Shared p_uploadedFiles As New List(Of String)
    Public Shared ReadOnly Property ImageArray() As List(Of String)
        Get
            Return p_imageArray
        End Get
    End Property
    Public Shared ReadOnly Property UploadedFiles() As List(Of String)
        Get
            Return p_uploadedFiles
        End Get
    End Property
    Public Property UploadFolder() As String
        Get
            Return p_uploadFolder
        End Get
        Set(ByVal value As String)
            p_uploadFolder = value
        End Set
    End Property
    '**********************************
    ' Initialize class
    '**********************************
    Public Sub New()
        ' Templater
        p_basePath = GetBasePath()
        p_mailTemplatePath = ""
        p_addHTMLBasePath = False
        p_parsePaths = False
        p_preMailerCss = False
        ' SMTP
        p_SmtpUserName = ""
        p_SmtpPassword = ""
        p_SmtpHost = ""
        p_SmtpEnableSsl = False
        p_SmtpPort = 25
        ' Mail
        p_emailRecipient = ""
        p_emailFrom = ""
        p_emailCC = ""
        p_emailBCC = ""
        p_emailReplyTo = ""
        p_eMailSubject = ""
        p_isBodyHtml = True
        p_emailPriority = "Normal"
        p_eMailEncoding = "utf-8"
        p_systemMailEncoding = Encoding.UTF8
        p_filesToAttach = Nothing
        p_attachmentFolder = "SmartMailerAttachments"
        p_attachmentMaxLength = 10000
        p_saveAttachments = True
        p_fileNamingStyle = ""
        ' Imaging Resize
        p_imageSizes = ""
        p_preventEnlarge = True
        p_preserveAspectRatio = True
        ' Crop
        p_cropSizes = ""
        p_cropPosition = "center-middle"
        ' Watermark
        p_watermarkMask = ""
        p_watermarkPadding = 5
        p_watermarkOpacity = 25
        p_watermarkSizes = "128, 128 | 64, 64, | 32, 32 |"
        p_watermarkAlign = "Center-Middle"
        ' Caption
        p_captionText = ""
        p_captionOpacity = 100
        p_captionFont = "Arial"
        p_captionFontSizes = "16, 14, 12"
        p_captionFontColor = "Black"
        p_captionFontStyle = "Regular"
        p_captionAlign = "Center-Middle"
        p_captionPadding = 5
        p_uploadFolder = "SmartMailerUploads"
    End Sub

    ' **********************************************************************************
    ' Section templating
    '***********************************************************************************
    ' Retrieve any page on internet as email body with HttpWebRequest & HttpWebResponse 
    Public Function GetMailBody() As String
        Dim mailBodyHTML As String = ""
        Dim serverStatusCode As Integer
        Dim HttpRequestURI As String = p_basePath & p_mailTemplatePath

        Try
            Dim objHttpHTML As HttpWebRequest
            Dim eMailTemplate As New Uri(HttpRequestURI)
            objHttpHTML = DirectCast(WebRequest.Create(eMailTemplate), HttpWebRequest)
            objHttpHTML.Timeout = 5 * 1000 ' 5 Seconds to avoid getting locked up

            ' Get response and check response code. 
            Dim serverResponse As HttpWebResponse = objHttpHTML.GetResponse()
            serverStatusCode = (serverResponse.StatusCode)

            If serverStatusCode <> 200 Then
                Throw New WebException("Email body creation failed. The server responded With HTTP status code: " & serverStatusCode.ToString)
                'If response code 200 OK and response length > 0, read response into mailBodyHTML.
            ElseIf (serverResponse.ContentLength > 0) Then
                Dim reader As New StreamReader(serverResponse.GetResponseStream())
                mailBodyHTML = reader.ReadToEnd().ToString
                reader.Close()
                serverResponse.Close()
            Else
                Throw New WebException("The email body is empty.")
            End If
        Catch e As WebException
            'Add errors to List(Of WebException), if any.
            ErrorCodes.Add(e)
        End Try

        ' If (p_addHTMLBasePath) call function to add <base href="p_basePath" /> tag to HTML <head>
        If (p_addHTMLBasePath) Then mailBodyHTML = AddBasePath(mailBodyHTML)
        ' If (p_parsePaths) call regular expression function to make relative paths absolute
        If (p_parsePaths) Then mailBodyHTML = MakeRelativeLinksAbsolute(mailBodyHTML)
        ' If (p_preMailerCss) pass mailBodyHTML to PreMailer to inline css.
        If (p_preMailerCss) Then mailBodyHTML = MoveHTMLStylesInline(mailBodyHTML)

        Return mailBodyHTML
    End Function

    ' Writes base path into the html document <head>, the quickest way to fix relative links.
    Public Function AddBasePath(ByVal MailBodyHTML As String) As String
        If InStr(MailBodyHTML, "<base href=") = 0 Then Replace(MailBodyHTML, "<head>", "<head>" & vbCrLf & "<base href=""" & p_basePath & """/>" & vbCrLf)
        Return MailBodyHTML
    End Function

    'Must have Premailer.Net assembly installed for inlining css.
    Public Function MoveHTMLStylesInline(ByVal MailBodyHTML As String) As String
        Try
            ' Create variable to hold Premailer result and initialize Premailer class.
            Dim strBodyHTML As New Net.PreMailer(MailBodyHTML, New Uri(p_basePath))
            Dim strResult As InlineResult = strBodyHTML.MoveCssInline(True)
            MailBodyHTML = strResult.Html
        Catch e As WebException
            'Add errors to List(Of WebException), if any.
            ErrorCodes.Add(e)
        End Try
        Return MailBodyHTML
    End Function

    ' Parse HTML with regular expression and make relative links absolute with p_basepath
    Public Function MakeRelativeLinksAbsolute(ByVal MailBodyHTML As String) As String
        ' Declare & intialize variables
        Dim strHTMLBody As String = MailBodyHTML
        ' Set regex variables 
        Dim strSubMatch As String = ""
        Dim strFullUrl As String = ""
        Dim Pattern As String = "[=""]\/?([^""\s]*(\.gif|\.jpg|\.jpeg|\.png|\.css|\.js|\.pdf|\.rtf|\.doc|\.xlsx|\.xls|\.exe))[""\s]"
        Dim Options As RegexOptions = RegexOptions.IgnoreCase
        Dim regex As Regex = New Regex(Pattern, Options)
        Dim regexSub As Regex = New Regex(Pattern, Options)
        Dim Matches As MatchCollection = regex.Matches(strHTMLBody)

        Try
            For Each Match As Match In Matches
                ' Filter out absolute links
                If InStr(Match.ToString, "://") = 0 And InStr(LCase(Match.ToString), "mailto:") = 0 And InStr(LCase(Match.ToString), "javascript:") = 0 Then
                    ' Remove the " marks at each end of string "
                    strSubMatch = regexSub.Replace(Match.ToString, "$1")
                    ' Concatenate the two strings into FullPath
                    strFullUrl = p_basePath & strSubMatch
                    ' Execute the replace
                    strHTMLBody = Replace(strHTMLBody, "/" & strSubMatch, strFullUrl)
                End If
            Next ' Go to next link
        Catch e As WebException
            'Add errors to List(Of WebException), if any.
            ErrorCodes.Add(e)
        End Try

        Return strHTMLBody 'MailBodyHTML
    End Function

    ' MakeAbsoluteUrl() converts relative URL to Absolute URI
    Public Function MakeAbsoluteUrl(ByVal RelativeUrl As String) As Uri
        ' get action tags, relative or absolute
        Dim URL As Uri = New Uri(RelativeUrl, UriKind.RelativeOrAbsolute)
        ' Make it absolute if it's relative
        If Not URL.IsAbsoluteUri Then
            Dim basePath As Uri = New Uri(p_basePath)
            URL = New Uri(basePath, URL)
        End If
        Return URL
    End Function

    ' GetBasePath() returns application base path
    Public Function GetBasePath() As String
        ' Get base path 
        Dim strPath As String = ""
        Dim strProtocol As String = Current.Request.Url.Scheme & "://"
        Dim port As String = Current.Request.Url.Port
        If port <> 80 And port <> 433 Then port = ":" & port Else port = ""
        Dim strAppInstallFolder As String = p_appInstallFolder
        If Not String.IsNullOrEmpty(strAppInstallFolder) Then strAppInstallFolder &= "/"
        strPath = strProtocol & Current.Request.ServerVariables("SERVER_NAME") & port & "/" & strAppInstallFolder
        Return strPath
    End Function

    '***********************************************************************************
    ' Section eMail
    '***********************************************************************************
    ' GetAttachements() processes uploaded attachments.
    Public Sub GetAttachments()
        ' Get upload files &  Call upload
        Getfiles()
        ConstructUpload()
        ' Process images
        ResizeImages()
        CropImages()
        AddWaterMark()
        AddTextCaption()
    End Sub
    ' If property SaveAttachments is false delete uploaded files.
    Public Sub DeleteAttachments()
        If p_emailAttachments.Count > 0 Then
            Try
                For i = 0 To p_emailAttachments.Count - 1
                    Dim filePath As String = p_emailAttachments(i).ToString.Trim
                    If File.Exists(filePath) Then
                        File.Delete(filePath)
                    End If
                Next
            Catch e As WebException
                'Add errors to List(Of WebException), if any.
                ErrorCodes.Add(e)
            End Try
        End If
    End Sub

    ' SendWebMail() uses System.Web.Helpers.
    Public Sub SendWebMail()
        ' Process uploaded attachments 
        GetAttachments()
        ' Get the mail body html string
        Dim mailBody As String = GetMailBody()

        Try
            ' Send the email message using System.Web.Helpers.WebMail.
            Send(
                to:=p_emailRecipient,
                from:=p_emailFrom,
                subject:=p_eMailSubject,
                body:=mailBody,
                cc:=p_emailCC,
                bcc:=p_emailBCC,
                isBodyHtml:=p_isBodyHtml,
                priority:=p_emailPriority,
                filesToAttach:=p_emailAttachments,
                contentEncoding:=p_eMailEncoding,
                additionalHeaders:=p_additionalHeaders,
                replyTo:=p_emailReplyTo
                )

            'If property SaveAttachments is false call delete.
            If Not p_saveAttachments Then DeleteAttachments()

        Catch e As WebException
            'Add errors to List(Of WebException), if any.
            ErrorCodes.Add(e)
        End Try

        ' Send the user to thank you page on success.
        If Not String.IsNullOrEmpty(p_successRedirect) And ErrorCodes.Count < 1 Then
            Current.Response.Redirect(p_successRedirect)
        End If
    End Sub

    ' Send System.Net.Mail().
    ' TODO: embedded images
    ' Dim mailBodyHtml As AlternateView = AlternateView.CreateAlternateViewFromString(mailBody, New Mime.ContentType("text/html"))
    ' Dim mailBodyText As AlternateView = AlternateView.CreateAlternateViewFromString(mailText, Nothing, "text/plain")
    '.AlternateViews.Add(mailBodyHtml)
    '.AlternateViews.Add(mailBodyText)
    Public Sub SendSystemMail()
        ' Process uploaded attachments
        GetAttachments()
        ' Declare variables
        Dim mailSMTP As New SmtpClient()
        Dim message As New MailMessage()
        Dim mailBody As String = GetMailBody()
        Dim mailText As String = ""
        Dim priority As MailPriority = MailPriority.Normal

        If LCase(p_emailPriority) = "low" Then
            priority = MailPriority.Low
        ElseIf LCase(p_emailPriority) = "high" Then
            priority = MailPriority.High
        End If

        Try
            With message
                ' System.Net.Mail.MailMesage settings 
                If InStr(p_emailFrom, ",") Then
                    ' eMail address and name
                    Dim fromArray As Array = Split(p_emailFrom, ",")
                    .From = New MailAddress(fromArray(0).ToString.Trim, fromArray(1).ToString.Trim)
                Else ' eMail address only
                    .From = New MailAddress(p_emailFrom.Trim)
                End If

                ' Multiple recipients
                If InStr(p_emailRecipient, "|") Then
                    If Right(p_emailRecipient.TrimEnd, 1) <> "|" Then p_emailRecipient &= "|"
                    Dim recipientsArray As Array = Split(p_emailRecipient, "|")
                    For i = 0 To UBound(recipientsArray) - 1
                        ' Recipient email address and name
                        If InStr(recipientsArray(i).ToString, ",") Then
                            Dim recipientArray As Array = Split(recipientsArray(i), ",")
                            .To.Add(New MailAddress(recipientArray(0).ToString.Trim, recipientArray(1).ToString.Trim))
                        Else ' Recipient email address only
                            .To.Add(New MailAddress(recipientsArray(i).ToString.Trim))
                        End If
                    Next ' Single recipient 
                ElseIf InStr(p_emailRecipient, ",") Then
                    ' eMail address and name
                    Dim recipientArray As Array = Split(p_emailRecipient, ",")
                    .To.Add(New MailAddress(recipientArray(0).ToString.Trim, recipientArray(1).ToString.Trim))
                Else ' eMail address only
                    .To.Add(New MailAddress(p_emailRecipient.Trim))
                End If

                ' Add CC address(es) from input string.
                If Not String.IsNullOrEmpty(p_emailCC) Then
                    ' Multiple recipients
                    If InStr(p_emailCC, "|") Then
                        If Right(p_emailCC.TrimEnd, 1) <> "|" Then p_emailCC &= "|"
                        Dim emailCCArray As Array = Split(p_emailCC, "|")
                        For i = 0 To UBound(emailCCArray) - 1
                            ' Recipient email address and name
                            If InStr(emailCCArray(i).ToString, ",") Then
                                Dim CCArray As Array = Split(emailCCArray(i), ",")
                                .CC.Add(New MailAddress(CCArray(0).ToString.Trim, CCArray(1).ToString.Trim))
                            Else ' Recipient email address only
                                .CC.Add(New MailAddress(emailCCArray(i).ToString.Trim))
                            End If
                        Next ' Single recipient 
                    ElseIf InStr(p_emailCC, ",") Then
                        ' eMail address and name
                        Dim CCArray As Array = Split(p_emailCC, ",")
                        .CC.Add(New MailAddress(CCArray(0).ToString.Trim, CCArray(1).ToString.Trim))
                    Else ' eMail address only
                        .CC.Add(New MailAddress(p_emailCC.Trim))
                    End If
                End If

                ' Add BCC address(es) from input string
                If Not String.IsNullOrEmpty(p_emailBCC) Then
                    ' Multiple recipients
                    If InStr(p_emailBCC, "|") Then
                        If Right(p_emailBCC.TrimEnd, 1) <> "|" Then p_emailBCC &= "|"
                        Dim emailBCCArray As Array = Split(p_emailBCC, "|")
                        For i = 0 To UBound(emailBCCArray) - 1
                            ' Recipient email address and name
                            If InStr(emailBCCArray(i).ToString, ",") Then
                                Dim BCCArray As Array = Split(emailBCCArray(i), ",")
                                .Bcc.Add(New MailAddress(BCCArray(0).ToString.Trim, BCCArray(1).ToString.Trim))
                            Else ' Recipient email address only
                                .Bcc.Add(New MailAddress(emailBCCArray(i).ToString.Trim))
                            End If
                        Next  ' Single recipient 
                    ElseIf InStr(p_emailBCC, ",") Then
                        ' eMail address and name
                        Dim BCCArray As Array = Split(p_emailBCC, ",")
                        .Bcc.Add(New MailAddress(BCCArray(0).ToString.Trim, BCCArray(1).ToString.Trim))
                    Else ' eMail address only
                        .Bcc.Add(New MailAddress(p_emailBCC.Trim))
                    End If
                End If

                .Subject = p_eMailSubject
                .Body = mailBody
                .BodyEncoding = p_systemMailEncoding
                .HeadersEncoding = p_systemMailEncoding
                .SubjectEncoding = p_systemMailEncoding
                .IsBodyHtml = p_isBodyHtml
                .Priority = priority

                ' Add headers from passed in NameValueCollection
                If p_systemMailHeaders.Count > 0 Then
                    For i = 0 To p_systemMailHeaders.Count - 1
                        .Headers.Add(p_systemMailHeaders.GetKey(i), p_systemMailHeaders.Get(i))
                    Next
                End If

                ' Add attachments p_emailAttachments
                If p_emailAttachments.Count > 0 Then
                    For i = 0 To p_emailAttachments.Count - 1
                        .Attachments.Add(New Attachment(p_emailAttachments(i)))
                    Next
                End If

                ' Add addresses to ReplyToList
                If Not String.IsNullOrEmpty(p_emailReplyTo) Then
                    ' Multiple recipients
                    If InStr(p_emailReplyTo, "|") Then
                        If Right(p_emailReplyTo.TrimEnd, 1) <> "|" Then p_emailReplyTo &= "|"
                        Dim replyToArray As Array = Split(p_emailReplyTo, "|")
                        For i = 0 To UBound(replyToArray) - 1
                            ' Recipient email address and name
                            If InStr(replyToArray(i).ToString, ",") Then
                                Dim replyTo As Array = Split(replyToArray(i), ",")
                                .ReplyToList.Add(New MailAddress(replyTo(0).ToString.Trim, replyTo(1).ToString.Trim))
                            Else ' Recipient email address only
                                .ReplyToList.Add(New MailAddress(replyToArray(i).ToString.Trim))
                            End If
                        Next  ' Single recipient 
                    ElseIf InStr(p_emailReplyTo, ",") Then
                        ' eMail address and name
                        Dim replyTo As Array = Split(p_emailReplyTo, ",")
                        .ReplyToList.Add(New MailAddress(replyTo(0).ToString.Trim, replyTo(1).ToString.Trim))
                    Else ' eMail address only
                        .ReplyToList.Add(New MailAddress(p_emailReplyTo.Trim))
                    End If
                End If
            End With

            With mailSMTP
                ' System.Net.SmtpClient settings
                .Host = p_SmtpHost
                .Credentials = New NetworkCredential(p_SmtpUserName, p_SmtpPassword)
                .EnableSsl = p_SmtpEnableSsl
                .Port = p_SmtpPort
                .Timeout = 5 * 1000
                ' Send the email
                .Send(message)
            End With

            ' Clean up.
            message.Dispose()
            'If property SaveAttachments is false call delete.
            If Not p_saveAttachments Then DeleteAttachments()

        Catch e As WebException
            'Add errors to List(Of WebException), if any.
            ErrorCodes.Add(e)
        End Try

        ' Send the user to thank you page on success.
        If Not String.IsNullOrEmpty(p_successRedirect) And ErrorCodes.Count < 1 Then
            Current.Response.Redirect(p_successRedirect)
        End If

    End Sub

    '***********************************************************************************
    ' Setion file uploads
    '***********************************************************************************
    ' ProcessUploads() constructs file upload and imaging without an email message.
    Public Sub ProcessUploads()
        ' Get upload files & call upload
        Getfiles()
        UploadFiles()
        ' Process images
        ResizeImages()
        CropImages()
        AddWaterMark()
        AddTextCaption()
    End Sub

    ' Getfiles() adds the HttpFileCollection to p_filesToAttach
    Public Sub Getfiles()
        p_filesToAttach = Current.Request.Files()
    End Sub

    ' UploadFiles() called by the file upload methods, uploads files & creates list of images for processing
    Public Overloads Sub UploadFiles(ByVal upFile As HttpPostedFile, ByVal FilePath As String)
        Dim fileType As String = Left(upFile.ContentType.ToLower(), 5)
        'Save the file 
        upFile.SaveAs(FilePath)
        ' Add Image filePath to List(Of p_imagesUploaded), a temporary list passed to imaging.
        If fileType.Equals("image") Then
            p_imagesUploaded.Add(FilePath)
        End If
    End Sub ' UploadFiles()

    ' ConstructUpload() called by mail subs. uploads files and creates list of email attachments.
    Public Sub ConstructUpload()
        ' Get full path to the attachments folder
        Dim FolderPath As String = Current.Request.MapPath(p_attachmentFolder) & "\"
        Dim httpFiles As HttpFileCollection = p_filesToAttach ' p_filesToAttach from Getfiles()

        Try
            For i = 0 To httpFiles.Count - 1
                Dim upFile As HttpPostedFile = httpFiles(i)
                Dim fileName As String = GetFileName(upFile.FileName)
                Dim upFileLength As Integer = upFile.ContentLength()
                If upFileLength > 0 Then
                    'verify file exists
                    If Not String.IsNullOrEmpty(fileName) Then
                        'set file paths
                        Dim filePath As String = Combine(FolderPath & GetSavedFileName(FolderPath, fileName))
                        'check file size
                        If upFileLength > (p_attachmentMaxLength * 1000) Then
                            Throw New WebException("Upload file:" & fileName & " exceeds the allowed size of: " & (p_attachmentMaxLength * 1000) & " kilobytes.")
                        Else
                            'Create the directory if it does not exist.
                            If Not Directory.Exists(FolderPath) Then
                                Directory.CreateDirectory(FolderPath)
                            End If
                            ' Delete existing files with same name before uploading new.
                            If File.Exists(filePath) Then
                                File.Delete(filePath)
                            End If
                            'Save the file 
                            UploadFiles(upFile, filePath)
                            ' Add the filePath to: List(Of p_emailAttachments)
                            p_emailAttachments.Add(filePath)
                        End If
                    Else
                        Throw New WebException("Upload file cannot be found.")
                    End If
                End If
            Next
        Catch e As WebException
            ErrorCodes.Add(e)
        End Try

    End Sub ' ConstructUpload

    ' UploadFiles() is accessed using ProcessUploads() sub.
    Public Overloads Sub UploadFiles()
        ' Get full path to the attachments folder
        Dim FolderPath As String = Current.Request.MapPath(p_uploadFolder) & "\"
        Dim httpFiles As HttpFileCollection = p_filesToAttach ' p_filesToAttach from Getfiles()

        Try
            For i = 0 To httpFiles.Count - 1
                Dim upFile As HttpPostedFile = httpFiles(i)
                Dim fileName As String = GetFileName(upFile.FileName)
                Dim upFileLength As Integer = upFile.ContentLength()
                If upFileLength > 0 Then
                    'verify file exists
                    If Not String.IsNullOrEmpty(fileName) Then
                        'set file paths
                        Dim filePath As String = Combine(FolderPath & GetSavedFileName(FolderPath, fileName))
                        'check file size
                        If upFileLength > (p_attachmentMaxLength * 1000) Then
                            Throw New WebException("Upload file:" & fileName & " exceeds the allowed size of: " & (p_attachmentMaxLength * 1000) & " kilobytes.")
                        Else
                            'Create the directory if it does not exist.
                            If Not Directory.Exists(FolderPath) Then
                                Directory.CreateDirectory(FolderPath)
                            End If
                            ' Delete existing files with same name before uploading new.
                            If File.Exists(filePath) Then
                                File.Delete(filePath)
                            End If
                            'Save the file 
                            UploadFiles(upFile, filePath)
                            p_uploadedFiles.Add(filePath)
                        End If
                    Else
                        Throw New WebException("Upload file cannot be found.")
                    End If
                End If
            Next
        Catch e As WebException
            ErrorCodes.Add(e)
        End Try
    End Sub ' UploadFiles()

    '***********************************************************************************
    ' Section WebImage imaging
    '***********************************************************************************
    ' ResizeImages() accepts multiple arguments as string in format: "width, height, suffix |"
    Public Sub ResizeImages()
        If Not String.IsNullOrEmpty(p_imageSizes) And p_imagesUploaded.Count > 0 Then
            Try
                ' Split input string into array of image specifications
                If Right(p_imageSizes.TrimEnd, 1) <> "|" Then p_imageSizes &= "|"
                Dim imageVarients As Array = Split(p_imageSizes, "|")
                For n = 0 To p_imagesUploaded.Count - 1
                    For i = 0 To UBound(imageVarients) - 1
                        ' Split imageVarients(i) into imageWidth, imageHeight & imageSuffix 
                        Dim imageSpecs As Array = Split(imageVarients(i), ",")
                        Dim imageWidth As Integer = Nothing
                        Dim imageHeight As Integer = Nothing
                        Dim imageSuffix As String = imageSpecs(2).ToString.Trim
                        If Not String.IsNullOrEmpty(imageSuffix) Then imageSuffix = "_" & imageSuffix
                        If IsNumeric(imageSpecs(0).ToString.Trim) Then imageWidth = CInt(imageSpecs(0).ToString.Trim)
                        If IsNumeric(imageSpecs(1).ToString.Trim) Then imageHeight = CInt(imageSpecs(1).ToString.Trim)

                        ' Get file name variables
                        Dim filePath As String = p_imagesUploaded(n)
                        Dim image As New WebImage(filePath)
                        Dim imageFileName As String = GetFileName(filePath)
                        Dim imageShortName As String = GetFileNameWithoutExtension(imageFileName)
                        Dim imageExtension As String = GetExtension(imageFileName)
                        Dim folderPath As String = GetDirectoryName(filePath) & "\"
                        Dim imageName As String = imageShortName & imageSuffix & imageExtension
                        Dim imagePath As String = Combine(folderPath & imageName)

                        ' If there's an image, resize it using imageSpecs
                        If imageWidth > 0 And imageHeight > 0 Then
                            If Not image Is Nothing Then
                                ' Save resized image
                                image.Resize(width:=imageWidth, height:=imageHeight, preserveAspectRatio:=p_preserveAspectRatio, preventEnlarge:=p_preventEnlarge)
                                image.Save(imagePath)
                                ' Add new image varient to list p_emailAttachments.
                                If Not String.IsNullOrEmpty(imageSuffix) Then
                                    p_emailAttachments.Add(imagePath)
                                End If
                                ' Add all image paths to p_imageArray for further processing.
                                p_imageArray.Add(imagePath)
                            Else
                                Throw New WebException("The image cannot be found.")
                            End If
                        Else
                            Throw New WebException("Resized image must be larger than 0 pixels in width and height.")
                        End If
                    Next
                Next
                ' Clear p_imagesUploaded temporary list of images after adding them to p_imageArray.
                p_imagesUploaded.Clear()
            Catch e As WebException
                ErrorCodes.Add(e)
            End Try

        End If
    End Sub

    ' CropImage() accepts size arguments as string in format: "width, height | width, height |"
    ' Location arguments are: left-top, left-middle, left-bottom, center-top, center-middle, center-bottom, right-top, right-middle, right-bottom.
    Public Sub CropImages()
        If Not String.IsNullOrEmpty(p_cropSizes) And p_imageArray.Count > 0 Then
            Try
                ' Get p_cropPosition from string and split into horizontal/vertical postitions 
                Dim cropHorizontal As String = ""
                Dim cropVertical As String = ""
                Dim cropPosition As Array = Split(p_cropPosition, "-")
                If Not String.IsNullOrEmpty(cropPosition(0).ToString.Trim) Then cropHorizontal = LCase(cropPosition(0).ToString.Trim)
                If Not String.IsNullOrEmpty(cropPosition(1).ToString.Trim) Then cropVertical = LCase(cropPosition(1).ToString.Trim)

                ' Split input string into array of image crop sizes
                Dim n As Integer = 0
                If Right(p_cropSizes.TrimEnd, 1) <> "|" Then p_cropSizes &= "|"
                Dim cropSizeVars As Array = Split(p_cropSizes, "|")
                ' Loop through p_imageArray in sync with p_cropSizes
                For i = 0 To p_imageArray.Count - 1
                    ' Split cropVarients(i) into imageWidth, imageHeight 
                    Dim cropSpecs As Array = Split(cropSizeVars(n), ",")
                    Dim cropWidth As Integer = Nothing
                    Dim cropHeight As Integer = Nothing
                    If IsNumeric(cropSpecs(0).ToString.Trim) Then cropWidth = CInt(cropSpecs(0).ToString.Trim)
                    If IsNumeric(cropSpecs(1).ToString.Trim) Then cropHeight = CInt(cropSpecs(1).ToString.Trim)

                    If cropWidth > 0 And cropHeight > 0 Then
                        ' Load image
                        Dim filePath As String = p_imageArray(n)
                        Dim image As New WebImage(filePath)
                        ' Get crop area arguments
                        Dim imageWidth As Integer = image.Width
                        Dim imageHeight As Integer = image.Height
                        Dim y_Top As Integer = Nothing
                        Dim x_Left As Integer = Nothing
                        Dim y_Bottom As Integer = Nothing
                        Dim x_Right As Integer = Nothing

                        ' X axis coordinates
                        Select Case cropHorizontal
                            Case "left"
                                x_Left = 0
                                x_Right = imageWidth - cropWidth
                                If x_Right < 0 Then x_Right = 0
                            Case "center"
                                x_Left = imageWidth - cropWidth
                                x_Left = Floor(x_Left / 2)
                                If x_Left < 0 Then x_Left = 0
                                x_Right = imageWidth - cropWidth
                                x_Right = Ceiling(x_Right / 2)
                                If x_Right < 0 Then x_Right = 0
                            Case "right"
                                x_Left = imageWidth - cropWidth
                                If x_Left < 0 Then x_Left = 0
                                x_Right = 0
                        End Select

                        ' Y axis coordinates
                        Select Case cropVertical
                            Case "top"
                                y_Top = 0
                                y_Bottom = imageHeight - cropHeight
                                If y_Bottom < 0 Then y_Bottom = 0
                            Case "middle"
                                y_Top = imageHeight - cropHeight
                                y_Top = Floor(y_Top / 2)
                                If y_Top < 0 Then y_Top = 0
                                y_Bottom = imageHeight - cropHeight
                                y_Bottom = Ceiling(y_Bottom / 2)
                                If y_Bottom < 0 Then y_Bottom = 0
                            Case "bottom"
                                y_Top = imageHeight - cropHeight
                                If y_Top < 0 Then y_Top = 0
                                y_Bottom = 0
                        End Select

                        ' If there's an image, crop it
                        If Not image Is Nothing Then
                            ' Crop the image
                            image.Crop(y_Top, x_Left, y_Bottom, x_Right)
                            ' Save cropped image (filePath)
                            image.Save(filePath)
                            ' Incrementing n first eliminates substraction of 1 from upper bound.
                            n += 1
                            ' Reset counter to 0 when cropSizeVars reaches upperbound. 
                            If n = UBound(cropSizeVars) Then n = 0
                        Else
                            Throw New WebException("The image cannot be found.")
                        End If
                    Else
                        Throw New WebException("Cropped image must be larger than 0 pixels in width and height.")
                    End If
                Next
            Catch e As WebException
                ErrorCodes.Add(e)
            End Try

        End If
    End Sub

    ' AddWaterMark() p_watermarkAlign As String in format horizontal-vertical
    ' AddWaterMark() p_watermarkSizes As String in format "width, height |".
    Public Sub AddWaterMark()
        If Not String.IsNullOrEmpty(p_watermarkMask) And p_imageArray.Count > 0 Then
            Try
                ' Parse string for watermark alignment
                Dim horizontalAlign As String = ""
                Dim verticalAlign As String = ""
                Dim watermarkAlign As Array = Split(p_watermarkAlign, "-")
                If Not String.IsNullOrEmpty(watermarkAlign(0).ToString.Trim) Then horizontalAlign = watermarkAlign(0).ToString.Trim
                If Not String.IsNullOrEmpty(watermarkAlign(1).ToString.Trim) Then verticalAlign = watermarkAlign(1).ToString.Trim

                ' Parse string for watermark size
                Dim n As Integer = 0
                If Right(p_watermarkSizes.TrimEnd, 1) <> "|" Then p_watermarkSizes &= "|"
                Dim watermarkSizes As Array = Split(p_watermarkSizes, "|")
                Dim WatermarkImage As New WebImage(Current.Request.MapPath(p_watermarkMask))

                For i = 0 To p_imageArray.Count - 1
                    Dim width As Integer = Nothing
                    Dim height As Integer = Nothing
                    Dim watermarkSize As Array = Split(watermarkSizes(n), ",")
                    If Not String.IsNullOrEmpty(watermarkSize(0).ToString.Trim) Then width = CInt(watermarkSize(0).ToString.Trim)
                    If Not String.IsNullOrEmpty(watermarkSize(1).ToString.Trim) Then height = CInt(watermarkSize(1).ToString.Trim)

                    Dim filePath As String = p_imageArray(i)
                    Dim image As New WebImage(filePath)
                    If Not WatermarkImage Is Nothing Then
                        If Not image Is Nothing Then
                            Dim imagePath As String = filePath
                            image.AddImageWatermark(WatermarkImage, width:=width, height:=height, horizontalAlign:=horizontalAlign, verticalAlign:=verticalAlign, opacity:=p_watermarkOpacity, padding:=p_watermarkPadding)
                            image.Save(imagePath)
                        Else
                            Throw New WebException("The image cannot be found.")
                        End If
                    Else
                        Throw New WebException("The watermark mask cannot be found.")
                    End If
                    ' Incrementing n first eliminates substraction of 1 from upper bound.
                    n += 1
                    ' Reset counter to 0 when watermarkSizes reaches upperbound. 
                    If n = UBound(watermarkSizes) Then n = 0
                Next
            Catch e As WebException
                ErrorCodes.Add(e)
            End Try
        End If
    End Sub

    ' AddTextCaption() 
    ' Property CaptionAlign As String in format "horizontal-vertical"
    ' Property CaptionFontSizes takes multiple font sizes in comma separated format: "16, 14, 12"  
    ' Property CaptionFontStyle valid values are: "Regular", "Bold", "Italic", "Underline", and "Strikeout".
    Public Sub AddTextCaption()
        If Not String.IsNullOrEmpty(p_captionText) And p_imageArray.Count > 0 Then
            Try
                ' Parse string for caption alignment
                Dim horizontalAlign As String = ""
                Dim verticalAlign As String = ""
                Dim captionAlign As Array = Split(p_captionAlign, "-")
                If Not String.IsNullOrEmpty(captionAlign(0).ToString.Trim) Then horizontalAlign = captionAlign(0).ToString.Trim
                If Not String.IsNullOrEmpty(captionAlign(1).ToString.Trim) Then verticalAlign = captionAlign(1).ToString.Trim

                Dim n As Integer = 0
                If Right(p_captionFontSizes.TrimEnd, 1) <> "," Then p_captionFontSizes &= ","
                Dim fontSizes As Array = Split(p_captionFontSizes, ",")
                For i = 0 To p_imageArray.Count - 1
                    Dim fontSize As Integer = fontSizes(n)
                    Dim filePath As String = p_imageArray(i)
                    Dim image As New WebImage(filePath)
                    Dim imagePath As String = filePath
                    If Not image Is Nothing Then
                        image.AddTextWatermark(p_captionText, fontSize:=fontSize, fontStyle:=p_captionFontStyle, fontColor:=p_captionFontColor, fontFamily:=p_captionFont, horizontalAlign:=horizontalAlign, verticalAlign:=verticalAlign, opacity:=p_captionOpacity, padding:=p_captionPadding)
                        image.Save(imagePath)
                    Else
                        Throw New WebException("The image cannot be found.")
                    End If
                    ' Incrementing n first eliminates need to substract 1 from upper bound.
                    n += 1
                    ' Reset counter to 0 when fontSizes reaches upperbound. 
                    If n = UBound(fontSizes) Then n = 0
                Next
            Catch e As WebException
                ErrorCodes.Add(e)
            End Try
        End If
    End Sub

    '***********************************************************************************
    ' Section file utilities
    '***********************************************************************************
    Public Function GetSavedFileName(ByVal FolderPath As String, ByVal HttpFileName As String) As String
        Dim fileName As String = ""

        Select Case p_fileNamingStyle
            Case "GUID" 'Save file with GUID System Unique Name
                fileName = MakeNameUnique(HttpFileName)
            Case "StoreUniqueName"
                'store file with unique name
                fileName = StoreUniqueName(FolderPath, HttpFileName)
            Case Else ' store file name as is
                fileName = GetFileName(HttpFileName)
        End Select

        Return fileName
    End Function 'Get File path

    ' Creates GUID names for images.
    Public Function MakeNameUnique(ByVal Filename As String) As String
        Dim g As Guid
        Dim n As String
        Dim MUN As String = ""
        g = Guid.NewGuid()
        n = GetExtension(Filename)
        MUN = g.ToString + n
        Return MUN
    End Function 'MakeNameUnique

    ' Extract filename and make it unique returns unique name with numeric suffix.
    Public Function StoreUniqueName(ByVal FolderPath As String, ByVal HttpFileName As String) As String
        Dim imgName As String = HttpFileName
        Dim imgFileName As String = GetFileNameWithoutExtension(HttpFileName)
        Dim imgFileExt As String = GetExtension(HttpFileName)

        Dim i As Integer = 0
        While File.Exists(FolderPath + imgName)
            imgName = imgFileName & "_" & i & imgFileExt
            i += 1
        End While

        Return imgName
    End Function
End Class
