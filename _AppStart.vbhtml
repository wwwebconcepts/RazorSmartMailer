@Code
    '=========================================================
    ' WWWeb Concepts wwwebconcepts.com
    ' James W. Threadgill james@wwwebconcepts.com
    ' Version 1.0.0.0 Copyright 2017
    '=========================================================
    Dim userDatabase As String = InitializeApplication.DataConn()
    Dim appInit As Boolean = HttpContext.Current.Application("AppInitialized")
    If Not (appInit) Then InitializeApplication.InitializeAppVars()
    WebSecurity.InitializeDatabaseConnection(userDatabase, "UserProfile", "UserId", "Email", autoCreateTables:=-1)

    ' To let users of this site log in using their accounts from other sites such as Microsoft, Facebook, and Twitter,
    ' you must update this site. For more information visit https://go.microsoft.com/fwlink/?LinkID=226949

    ' OAuthWebSecurity.RegisterMicrosoftClient(
    '     clientId:="",
    '     clientSecret:="")

    ' OAuthWebSecurity.RegisterTwitterClient(
    '     consumerKey:="",
    '     consumerSecret:="")

    ' OAuthWebSecurity.RegisterFacebookClient(
    '     appId:="",
    '     appSecret:="")

    ' OAuthWebSecurity.RegisterGoogleClient()

    WebMail.SmtpServer = InitializeApplication.SmtpServer()  'ConfigurationManager.AppSettings("SmtpServer")
    WebMail.EnableSsl = InitializeApplication.EnableSsl()    'ConfigurationManager.AppSettings("EnableSsl")
    WebMail.UserName = InitializeApplication.UserName()      'ConfigurationManager.AppSettings("UserName")
    WebMail.Password = InitializeApplication.Password()      'ConfigurationManager.AppSettings("Password")
    WebMail.From = InitializeApplication.From()              'ConfigurationManager.AppSettings("From")
    WebMail.SmtpPort = 25
    ' To learn how to optimize scripts and stylesheets in your site go to  https://go.microsoft.com/fwlink/?LinkID=248974
End Code