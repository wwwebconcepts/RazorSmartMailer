@Code
    ' Handle user request to swap theme
    If Not Request.QueryString("Theme") Is Nothing Then
        Dim swapTo As String = Request.QueryString("Theme")
        InitializeApplication.SwapTheme(swapTo)
    End If

    Layout = Helpers.SetTheme().ToString
    PageData("Title") = HttpContext.Current.Application("ContactFormHdr")

    ' Initialize page variables
    Dim human As Boolean
    Dim errormsg As String = ""
    Dim name As String = ""
    Dim phone As String = ""
    Dim city As String = ""
    Dim state As String = ""
    Dim zipcode As String = ""
    Dim email As String = ""
    Dim address As String = ""
    Dim message As String = ""
    Dim subject As String = ""
    Dim recipient As String = HttpContext.Current.Application("RecipientEmail")
    Dim recipientName As String = HttpContext.Current.Application("RecipientEmailName")


    ' Get properties from InitializeApplication
    Dim userDatabase As String = InitializeApplication.DataConn()
    Dim appInstallPath As String = InitializeApplication.AppInstallFolder()
    Dim recaptchaPublicKey As String = InitializeApplication.RecaptchaPublicKey()
    Dim recaptchaPrivateKey As String = InitializeApplication.RecaptchaPrivateKey()
    Dim eMailContentEncoding As String = InitializeApplication.eMailContentEncoding()

    ' Initialize ReCAPTCHA class & declare variables
    Dim theCAPTCHA As New ReCAPTCHANet
    Dim theme As String = HttpContext.Current.Application("ReCAPTCHATheme")
    Dim language As String = HttpContext.Current.Application("ReCAPTCHALanguage")

    ' Configure Recaptcha properties
    With theCAPTCHA
        .PublicKey = "" & recaptchaPublicKey & ""
        .PrivateKey = "" & recaptchaPrivateKey & ""
        .QueryParameter = ""
        .RedirectURL = ""
        .FailURL = ""
    End With

    ' Setup validation variables
    Dim phonepattern As String = "\(\d{3}\)\s\d{3}-\d{4}( x\d{4})?|x\d{4}"
    Dim zipcodepattern As String = "\d{5}-?(\d{4})?$"
    Dim emailpattern As String = "^[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,4}$"
    Dim applicationError As String = ("<div class=""errortext"">The site currently has no email address configured for the forms. You must add an email address to the site configuration to use the forms.</div>")

    ' Setup validation
    Validation.RequireField("Name", "Name cannot be blank.<br />")
    Validation.RequireField("Email", "Email address cannot be blank.<br />")
    Validation.Add("Email", Validator.Regex(emailpattern, "Invalid email address.<br />"))
    Validation.RequireField("Address", "Address cannot be blank.<br />")
    Validation.RequireField("Phone", "Phone cannot be blank.<br />")
    Validation.Add("Phone", Validator.Regex(phonepattern, "Invalid phone number. The required pattern is (___) ___-____ x____, extension being optional.<br />"))
    Validation.RequireField("City", "City cannot be blank.<br />")
    Validation.RequireField("State", "State cannot be blank.<br />")
    Validation.RequireField("ZipCode", "Zip Code cannot be blank.<br />")
    Validation.Add("ZipCode", Validator.Regex(zipcodepattern, "Invalid zip code. Allowed patterns are: _____-____ or _____.<br />"))
    Validation.RequireField("Recaptcha", "Recaptcha response cannot be blank.<br />")
    Validation.RequireField("Message", "Message cannot be blank.<br />")
    Validation.Add("Message", Validator.StringLength(maxLength:=500, minLength:=25, errorMessage:="Message must be at least 25 and no more than 500 characters.<br />"))

    If IsPost Then
        AntiForgery.Validate()
        name = Request.Form("Name")
        address = Request.Form("Address")
        email = Request.Form("Email")
        phone = Request.Form("Phone")
        city = Request.Form("City")
        state = Request.Form("State")
        zipcode = Request.Form("ZipCode")
        message = Request.Form("Message")
        subject = Request.Form("_subject")
        human = theCAPTCHA.Construct()

        If Not human Then
            human = True 'ModelState.AddFormError("Recaptcha response was not correct.")
        End If

        If Validation.IsValid() Then
            ' If successful, set templatePath and redirect to thanks page on success.
            Dim templatePath As String = "Mail_Templates/contact_template.vbhtml?N=" & Server.UrlEncode(name).Trim & "&A=" & Server.UrlEncode(address).Trim & "&E=" & Server.UrlEncode(email).Trim & "&P=" & Server.UrlEncode(phone).Trim & "&C=" & Server.UrlEncode(city).Trim & "&S=" & Server.UrlEncode(state).Trim & "&Z=" & Server.UrlEncode(zipcode).Trim & "&M=" & Server.UrlEncode(message).Trim & ""
            Dim thanksPath As String = ("Thanks?N=" & Server.UrlEncode(name).Trim & "&A=" & Server.UrlEncode(address).Trim & "&E=" & Server.UrlEncode(email).Trim & "&P=" & Server.UrlEncode(phone).Trim & "&C=" & Server.UrlEncode(city).Trim & "&S=" & Server.UrlEncode(state).Trim & "&Z=" & Server.UrlEncode(zipcode).Trim & "&M=" & Server.UrlEncode(message).Trim & "")
            Try
                ' Initialize RazorSmartMailer 
                Dim theMailer As New RazorSmartMailer

                ' Configure RazorSmartMailer  properties
                With theMailer
                    ' templater properties
                    .AppInstallFolder = "" & appInstallPath & ""
                    .MailTemplatePath = "" & templatePath & ""
                    .AddHTMLBasePath = False
                    .ParsePaths = True
                    .PreMailerCss = True
                    ' SMTP Server for System.Net.Mail
                    .SmtpUserName = "james@wwwebconcepts"
                    .SmtpPassword = "Karen7065"
                    .SmtpHost = "192.168.1.16"
                    ' email properties
                    .SuccessRedirect = "" & thanksPath & ""
                    .AttachmentFolder = "UploadedAttachments"
                    .eMailRecipient = "" & recipient & "," & recipientName & ""
                    .eMailFrom = "" & email & "," & name & ""
                    .eMailSubject = "" & subject & ""
                    .eMailCC = "" & email & "," & name & ""
                    ' Imaging Resize
                    .ImageSizes = "525, 525, large | 175, 175, thumb | 325, 325,"
                    ' Crop
                    .CropSizes = "500, 500 | 150, 150 | 300, 300"
                    .CropPosition = "Center-Middle"
                    ' Watermark
                    .WatermarkMask = "~/Masks/logo_www.png"
                    .WatermarkAlign = "Center-Middle"
                    .WatermarkSizes = "128, 128 | 32, 32 | 64, 64"
                    .WatermarkPadding = 10
                    .WatermarkOpacity = 25
                    ' Caption
                    .CaptionText = "http://wwwebconcepts.com"
                    .CaptionFontSizes = "22, 10, 12"
                    .CaptionAlign = "Center-Top"
                End With

                ' If Recaptcha has errors ModelState.AddFormError()
                If theCAPTCHA.ErrorCodes.Count > 0 Then
                    For i = 0 To theCAPTCHA.ErrorCodes.Count - 1
                        ModelState.AddFormError(theCAPTCHA.ErrorCodes(i).Message & ": " & theCAPTCHA.ErrorCodes(i).HelpLink)
                    Next
                Else
                    ' Send mail
                    theMailer.SendSystemMail()
                    ' If RazorSmartMailer has errors ModelState.AddFormError()
                    If theMailer.ErrorCodes.Count > 0 Then
                        For i = 0 To theMailer.ErrorCodes.Count - 1
                            ModelState.AddFormError(theMailer.ErrorCodes(i).Message & ": " & theMailer.ErrorCodes(i).HelpLink)
                        Next
                    End If
                End If
            Catch e As Exception
                ModelState.AddFormError(e.Message)
            End Try
        End If
    End If

    ' Retrieve the dataset rows for the locations drop down
    Dim datasource As Database = Database.Open(userDatabase)
    Dim sqlQuery As String = "SELECT D_Destination FROM webpages_Destinations Order By D_Destination Asc"
    Dim rs_locations = datasource.Query(sqlQuery)
End Code

<hgroup class="title">
    <h1>@PageData("Title").</h1>
    <h2>@HttpContext.Current.Application("ContactFormSubHdr")</h2>
</hgroup>

<div Class="row">
    <div class="col-md-12">
        <div id="form-container">
            <div class="forms">
                <form method="post" enctype="multipart/form-data">
                    @AntiForgery.GetHtml()
                    @* If at least one validation error exists, notify the user *@
                    @Html.ValidationSummary("Please correct the errors and try again.", excludeFieldErrors:=True, htmlAttributes:=Nothing)
                    @If String.IsNullOrEmpty(recipient) Then
                    @Html.Raw(applicationError)
                    Else
                    @<fieldset>
                        <legend class="notice-text">
                            @HttpContext.Current.Application("FormInstructionsLbl")
                        </legend>
                        <div class="name">
                            <div class="form_lbl">
                                <Label for="Name">@HttpContext.Current.Application("FormNameLbl")</Label>
                            </div>
                            <div class="form_field">
                                @Html.ValidationMessage("Name")
                                <input name="Name" id="Name" type="text" class="long_txt" value="@name" @Validation.For("Name") maxlength="100" />
                            </div>
                        </div>
                        <div class="email">
                            <div class="form_lbl"><label for="Email">@HttpContext.Current.Application("FormEmailLbl")</label></div>
                            <div class="form_field">
                                @Html.ValidationMessage("Email")
                                <input name="Email" id="Email" type="text" class="long_txt" value="@email" @Validation.For("Email") maxlength="100" />
                            </div>
                        </div>
                        <div class="address">
                            <div class="form_lbl"><label for="Address">@HttpContext.Current.Application("FormAddressLbl")</label></div>
                            <div class="form_field">
                                @Html.ValidationMessage("Address")
                                <input name="Address" id="Address" type="text" class="long_txt" value="@address" @Validation.For("Address") maxlength="100" />
                            </div>
                        </div>
                        <div class="city">
                            <div class="form_lbl"><label for="City">@HttpContext.Current.Application("FormCityLbl")</label></div>
                            <div class="form_field">
                                @Html.ValidationMessage("City")
                                <input name="City" id="City" type="text" class="long_txt" value="@city" @Validation.For("City") maxlength="100" />
                            </div>
                        </div>
                        <div class="state">
                            <div class="form_lbl"><label for="State">@HttpContext.Current.Application("FormStateLbl")</label></div>
                            <div class="form_field">
                                @Html.ValidationMessage("State")
                                <select name="State" id="State" @Validation.For("State")>
                                    <option value="" @IIf(String.IsNullOrEmpty(state), Html.Raw("selected=""selected"""), "")>@HttpContext.Current.Application("FormStateLbl")</option>
                                    @For Each item In rs_locations
                                        @<option value="@item("D_Destination")" @IIf(item("D_Destination").ToString = state, Html.Raw("selected=""selected"""), "")>@item("D_Destination")</option>
                                    Next
                                </select>
                            </div>
                        </div>
                        <div class="zipcode">
                            <div class="form_lbl"><label for="ZipCode">@HttpContext.Current.Application("FormZipCodeLbl")</label></div>
                            <div class="form_field">
                                @Html.ValidationMessage("ZipCode")
                                <input name="ZipCode" id="ZipCode" class="zip" type="text" value="@zipcode" @Validation.For("ZipCode") maxlength="10" />
                            </div>
                        </div>
                        <div class="phone">
                            <div class="form_lbl"><label for="Phone">@HttpContext.Current.Application("FormPhoneLbl")</label></div>
                            <div class="form_field">
                                @Html.ValidationMessage("Phone")
                                <input name="Phone" id="Phone" type="text" value="@phone" @Validation.For("Phone") maxlength="20" />
                            </div>
                        </div>
                        <div class="File">
                            <div class="form_lbl"><label for="File">@HttpContext.Current.Application("FormFileLbl")</label></div>
                            <div class="form_field">
                                <input type="file" name="FileUpload1" id="FileUpload1" value="" />
                            </div>
                        </div>
                        <div class="message">
                            <div class="form_lbl"><label for="message">@HttpContext.Current.Application("FormMessageLbl")</label></div>
                            <div class="form_field">
                                @Html.ValidationMessage("Message")
                                <textarea name="Message" id="Message" cols="50" rows="6" @Validation.For("Message")>@message</textarea>
                            </div>
                        </div>
                        <div class="captcha">
                            <div class="form_lbl"><label for="recaptcha">@HttpContext.Current.Application("ReCaptchaLbl")</label></div><div class="form_field">@Html.ValidationMessage("Recaptcha") @Html.Raw(theCAPTCHA.GetControl(theme, language))</div>
                        </div>
                        <div Class="form_btn">
                            <input name="_Submit" id="_Submit" type="submit" class="button" value="@HttpContext.Current.Application("FormBtnTxt")" />
                        </div>
                        <input name="_subject" id="_subject" type="hidden" value="@HttpContext.Current.Application("ContactFormSubject")" />
                    </fieldset>
                    End If
                </form>
            </div>
        </div>
    </div>
</div>

@Section Scripts
    <script src="~/Scripts/jquery.validate.min.js"></script>
    <script src="~/Scripts/jquery.validate.unobtrusive.min.js"></script>
End Section