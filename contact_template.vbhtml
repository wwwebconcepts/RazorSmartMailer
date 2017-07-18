@Code
    ' Set page properties
    Layout = HttpContext.Current.Application("EmailTheme")
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
<div id="form-container">
    <div class="forms">
        <table>
            @If RazorSmartMailer.EmbeddedAttachments.Count > 0 Then
                @<tr>
                    <td Class="label"><div class="thanks_lbl">@HttpContext.Current.Application("eMailAttachments")</div></td>
                    <td>
                        <div class="thanks_field">
                            @For i = 0 To RazorSmartMailer.EmbeddedAttachments.Count - 1
                                Dim contentID As String = RazorSmartMailer.EmbeddedAttachments(i)
                                @<div class="embedded-image"><img src="cid:@contentID" /></div>
                            Next
                        </div>
                    </td>
                </tr>
                RazorSmartMailer.EmbeddedAttachments.Clear()
            End If
            <tr>
                <td class="label"><div class="thanks_lbl">@HttpContext.Current.Application("FormNameLbl")</div></td>
                <td><div class="thanks_field">@name</div></td>
            </tr>
            <tr>
                <td><div class="thanks_lbl">@HttpContext.Current.Application("FormEmailLbl")</div></td>
                <td><div class="thanks_field">@email</div></td>
            </tr>
            <tr>
                <td><div class="thanks_lbl">@HttpContext.Current.Application("FormAddressLbl")</div></td>
                <td><div class="thanks_field">@address</div></td>
            </tr>
            <tr>
                <td><div class="thanks_lbl">@HttpContext.Current.Application("FormPhoneLbl")</div></td>
                <td><div class="thanks_field">@phone</div></td>
            </tr>
            <tr>
                <td><div class="thanks_lbl">@HttpContext.Current.Application("FormCityLbl"), @HttpContext.Current.Application("FormStateLbl") @HttpContext.Current.Application("FormZipCodeLbl")</div></td>
                <td><div class="thanks_field">@city, @state @zipcode</div></td>
            </tr>
            <tr>
                <td><div class="thanks_lbl">@HttpContext.Current.Application("FormMessageLbl")</div></td>
                <td><div class="thanks_field">@message</div></td>
            </tr>
        </table>
    </div>
</div>
