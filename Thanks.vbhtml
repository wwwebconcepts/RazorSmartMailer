@Code
    ' Handle user request to swap theme
    If Not Request.QueryString("Theme") Is Nothing Then
        Dim swapTo As String = Request.QueryString("Theme")
        InitializeApplication.SwapTheme(swapTo)
    End If

    ' set page theme and title properties
    Layout = Helpers.SetTheme().ToString
    PageData("Title") = HttpContext.Current.Application("ContactThanksHdr")

    ' Initialize page variables
    Dim name As String = Request.QueryString("N")
    Dim email As String = Request.QueryString("E")
    Dim address As String = Request.QueryString("A")
    Dim phone As String = Request.QueryString("P")
    Dim city As String = Request.QueryString("C")
    Dim state As String = Request.QueryString("S")
    Dim zipcode As String = Request.QueryString("Z")
    Dim message As String = Request.QueryString("M")
End Code

<hgroup class="title">
    <h1>@PageData("Title")</h1>
    <h2>@HttpContext.Current.Application("ContactThanksSubHdr")</h2>
</hgroup>

<div Class="row">
    <div class="col-md-12">
        <div id="form-container">
            <div class="forms">
                @If RazorSmartMailer.eMailAttachments.Count > 0 Then
                    @<div Class="files">
                        <div Class="thanks_lbl">@HttpContext.Current.Application("eMailAttachments")</div>
                        <div class="thanks_field">
                            @For i = 0 To RazorSmartMailer.eMailAttachments.Count - 1
                                Dim fileName As String = RazorSmartMailer.eMailAttachments(i).ToString()
                                @Html.Raw(Path.GetFileName(fileName) & "<br />")
                            Next
                        </div>
                    </div>
                    RazorSmartMailer.eMailAttachments.Clear()
                End If
                <div class="name">
                    <div class="thanks_lbl">@HttpContext.Current.Application("FormNameLbl")</div>
                    <div class="thanks_field">@name</div>
                </div>
                <div class="address">
                    <div class="thanks_lbl">@HttpContext.Current.Application("FormAddressLbl")</div>
                    <div class="thanks_field">@address</div>
                </div>
                <div class="email">
                    <div class="thanks_lbl">@HttpContext.Current.Application("FormEmailLbl")</div>
                    <div class="thanks_field">@email</div>
                </div>
                <div class="phone">
                    <div class="thanks_lbl">@HttpContext.Current.Application("FormPhoneLbl")</div>
                    <div class="thanks_field">@phone</div>
                </div>
                <div class="city">
                    <div class="thanks_lbl">@HttpContext.Current.Application("FormCityLbl"), @HttpContext.Current.Application("FormStateLbl") @HttpContext.Current.Application("FormZipCodeLbl")</div>
                    <div class="thanks_field">@city, @state @zipcode</div>
                </div>
                <div class="message">
                    <div class="thanks_lbl">@HttpContext.Current.Application("FormMessageLbl")</div>
                    <div class="thanks_field">@message</div>
                </div>
            </div>
        </div>
    </div>
</div>
