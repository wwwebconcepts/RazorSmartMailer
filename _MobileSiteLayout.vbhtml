<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="utf-8" />
    <title>@PageData("Title")</title>
    @RenderSection("Meta", required:=False)
    <meta name="viewport" content="width=device-width, initial-scale=1.0" />
    <base href="@Helpers.GetBasePath()" />
    <link href="@Helpers.Canonical()" rel="canonical" />
    <link href="Content/themes/base/jquery.ui.all.css" rel="stylesheet" type="text/css" />
    <link href="Content/Site.css" rel="stylesheet" type="text/css" />
    <link href="Content/ApplicationStyle.css" rel="stylesheet" type="text/css" />
    <link href="bootstrap/css/bootstrap.css" rel="stylesheet" type="text/css" />
    <link href="favicon.ico" rel="shortcut icon" type="image/x-icon" />
    <script src="Scripts/jquery-3.1.1.js"></script>
    <script src="Scripts/modernizr-2.6.2.js"></script>
    <script src="Scripts/razor-portfolio.js"></script>
    <script src="bootstrap/js/npm.js"></script>
    @RenderSection("HeadScripts", required:=False)
</head>
<body>
    <div class="container">
        <header>
            <div class="content-wrapper">
                <div class="float-left">
                    <p class="site-title"><a href="">mobile</a></p>
                </div>
                <div class="float-right">
                    <section id="login">
                        @If WebSecurity.IsAuthenticated Then
                            @<text>
                                Hello, <a class="email" href="Account/Manage" title="Manage">@WebSecurity.CurrentUserName</a>!
                                <form id="logoutForm" action="Account/Logout" method="post">
                                    @AntiForgery.GetHtml()
                                    <a href="javascript:document.getElementById('logoutForm').submit()">Log out</a>
                                </form>
                            </text>
                        Else
                            @<ul>
                                <li><a href="Account/Register">Register</a></li>
                                <li><a href="Account/Login">Log in</a></li>
                            </ul>
                        End If
                    </section>
                    <nav>
                        <ul id="menu">
                            <li><a href="">Home</a></li>
                            <li><a href="About">About</a></li>
                            <li><a href="Contact">Contact</a></li>
                            <li><a href="Search">Search</a></li>
                        </ul>
                    </nav>
                </div>
            </div>
        </header>
        <div id="body">
            @RenderSection("featured", required:=False)
            <section class="content-wrapper main-content clear-fix">
                @RenderBody()
            </section>
        </div>
        <footer>
            <div class="content-wrapper">
                <div class="float-right"><a href="@Helpers.PageAsDesktop()" rel="nofollow">View Desktop Version</a></div>
                <a href="http://wwwebconcepts.com" target="_blank"><img src="icons/favicon.png" alt="WWWeb Concepts" width="32" height="32" class="www_logo" longdesc="http://wwwebconcepts.com" /></a>Razor Portfolio&#8482; &copy; @DateTime.Now.Year<br />
                <a href="http://wwwebconcepts.com" title="WWWeb Concepts Web Design, Development, SEO" class="footer">WWWeb Concepts | Web Design, Development, SEO</a>
            </div>
        </footer>
        @RenderSection("Scripts", required:=False)
    </div>
</body>
</html>