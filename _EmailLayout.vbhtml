<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="utf-8" />
    <base href="@Helpers.GetBasePath()" />
    @RenderSection("Meta", required:=False)
    <link href="~/Mail_Templates/css/mail_template_style.css" rel="stylesheet" type="text/css" />
    @RenderSection("HeadScripts", required:=False)
</head>
<body>
    <div class="container">
        <div class="mail_container">
            <div class="mail-content">
                <div class="header">
                    <img src="cid:mailbanner" alt="header" class="header_image" />
                    <!--<img src="~/Mail_Templates/MailTemplateImages/mailbanner.gif" alt="header" class="header_image" />-->
                </div>
                <div id="body">
                    @RenderSection("featured", required:=False)
                    <section class="content-wrapper main-content clear-fix">
                        @RenderBody()
                    </section>
                </div>
                <div class="clear"></div>
                <footer class="footer">
                    <a href="http://wwwebconcepts.com" target="_blank"><img src="cid:@RazorSmartMailer.EmbeddedImages(1)" alt="WWWeb Concepts" width="32" height="32" class="logo-image" longdesc="http://wwwebconcepts.com" /></a>Razor Portfolio&#8482; &copy; @DateTime.Now.Year
                    <br />
                    <a href="http://wwwebconcepts.com" title="WWWeb Concepts Web Design, Development, SEO">WWWeb Concepts | Web Design, Development, SEO</a>
                </footer>
                @RenderSection("Scripts", required:=False)
            </div>
        </div>
    </div>
</body>
</html>
@code
    RazorSmartMailer.EmbeddedImages.Clear()
End Code