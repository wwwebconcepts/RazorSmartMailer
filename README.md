# RazorSmartMailer 

VB.NET Library for sending email via  web forms. RazorSmartMailer has advanced HTML email templating and email messaging. RazorSmartMailer supports Helpers.WebMail, System.NET.Mail, attachment uploads and embed, embed linked resources, and image processing including resize, crop, watermark, and add text. RazorSmartMailer also supports inline CSS via the PreMailer.Net assembly. Razor Smart Mailer is your complete email solution.

## Dependencies
RazorSmartMailer requires the System.Web.Helpers and PreMailer assemblies. Property PreMailerCss set True moves Css inline via Premailer.Net. The image processing features implement the Web.Helpers.WebImage class. 

## Order of Imaging Operations
Resize, Crop, AddWaterMark,  AddCaption. Resize creates a list for the remaining methods.  RazorSmartMailer supports multiple operations on images in the order given.

## Usage
RazorSmartMailer returns five List(of String):
```vbnet
        ImageArray
        UploadedFiles
        EmailAttachments
        EmbeddedImages  
        EmbeddedAttachments
 ```
 The first three lists return the full system path to the named file collection. The Embedded lists contain the content IDs of the embedded files.

 ImageArray returns only images that have been resized and any other subsequent image processing. 

 Uploaded files returns the list of files uploaded.

 EmailAttachments returns the list of attachments including any image varients from imaging.

 EmbeddedImages retuns the list of linked resource embedded images.

 Embedded attachments returns the list of attachments embedded in the email body, including any image varients from imaging.

RazorSmartMailer has one more List(of WebException): ErrorCodes. This list returns any application errors.

 The RazorSmartMailer class has a number of properties you will need to set. Below are the properties and default settings as well as input formats.

```vbnet
'RazorSmartMailer calling code
Dim theMailer As New RazorSmartMailer

With theMailer
    'RazorSmartMailer templater properties
    .AppInstallFolder = "" ' Only if the application is not in the website root do you need to supply the folderpath.
    .MailTemplatePath = "" ' Path to email body template. May be an executable page in any ASP.NET coding language.
    .SuccessRedirect = "" ' Path to form confirmation "Thanks" page. 
    .PreMailerCss = False ' To Move CSS styles inline with Premailer.Net set True. 
    .AddHTMLBasePath = False '  When True writes base path into the html document <head>, the quickest way to fix relative links.
    .ParsePaths = False ' Parse HTML with regular expression and make relative links absolute.
 
    'SMTP Server properites for System.Net.Mail
    .SmtpUsername = ""
    .SmtpPassword = ""
    .SmtpHost = ""
    .SmtpEnableSsl = False
    .SmtpPort = 25

    'RazorSmartMailer sendmail properties
    .EmailFrom = "" ' From as string in format "emailaddress, name" or "emailaddress".
    .EmailRecipient = "" ' Multiple recipient lists are strings in format "emailaddress, name | emailaddress, name" or "emailaddress | emailaddress". Single addresses use the same format as From. The closing "|" may be omitted.
    .EmailCC = ""
    .EmailBC = ""
    .EmailReplyTo = ""
    .EMailEncoding = "utf-8"
    .EMailSubject = ""
    .IsBodyHtml = True
    .EmailPriority = "Normal"
    .AdditionalHeaders = Nothing
    .AttachmentFolder = "SmartMailerAttachments"
    .SaveAttachments = True
    SendWebMail() ' Set WebMail SMTP properties in _AppStart. Supports attachments. Does not support embedded images.

    'System.Net.Mail Send Properties
    .SystemMailHeaders = Nothing
    .SystemMailEncoding = Encoding.UTF8
    .EmbedAttachments = False
    .EmbedAllowExtensions = ".jpg, .jpeg, .gif, .png, .ico" ' These are the default allowed extensions for embedded attachments. 
    .ImagesToEmbed = "" ' String Format:  "image.gif, image.jpg, image.jpeg, image.ico, image.png"
    SendSystemMail() ' Send mail message using System.Net.Mail.

    'Resize
    .ImageSizes = "" ' String Format:  "width, height, suffix | width, height, suffix |". ' Resize only works if sizes defined. 
    .PreventEnlarge = True
    .PreserveAspectRatio = True

    'Crop
    .CropSizes = "" ' String Format:  "width, height | width, height". ' *Must match resize order. Crop only works if sizes defined in Resize and Crop.
    .CropPosition = "center-middle"

     'Watermark
     .WatermarkMask = "" ' Add path to watermark mask to enable watermarks.
     .WatermarkPadding = 10
     .WatermarkOpacity = 50
     .WatermarkSizes = "128, 128"
     .WatermarkAlign = "Center-Middle"

     'Captions
     .CaptionText = "" ' Add text to enable captions.
     .CaptionFont = "Ariel"
i     .CaptionFontSizes = "16"
     .CaptionFontColor = "Black"
     .CaptionFontStyle =  "Bold" Valid values are: "Regular", "Bold", "Italic", "Underline", and "Strikeout".
     .CaptionOpacity = 100
     .CaptionPadding = 10
     .CaptionAlign = "Center-Middle"
      'Uploader
     .UploadFolder = "SmartMailerUploads" 
     ProcessUploads() ' Constructs file upload and imaging without an email message.
End With
 ```

### Imaging with `RazorSmartMailer`
Below are the imaging properties. These properties are used with the email utility and the file upload utility. All images follow the same path through the methods: Resize, Crop, AddWaterMark,  AddCaption.  

Resize is the first image processing executed. It creates the ImageArray list of paths which all the succeeding imaging methods use.
We recommend the order from the largest to smallest, always saving any changes to original uploaded image for last.

 Each set of Resize instructions has three elements width, height, suffix. The two size elements are followed by a comma and each set is closed with a "|". To save an image with the original file name, leave the suffix element blank. You can also omit the closing "|" at the end of your string as the application will add it if not present. 


Let's deconstruct this input string: 
```vbnet 
"525, 525, large | 175, 175, thumb | 325, 325,"
 ```
 It makes three images, the first is 525 x 525 and is saved with the suffix "_large"; the second  image is 175 x 175 and is saved with the suffix  "_thumb"; the final image is 325 x 325 and saved with the original file name. 
 
 Why suffix naming? It provides a consistent method for naming generated images so one can easily display them. (Image name strings can be easily manipulated using the System.IO.Path class.) And it keeps all the images from an original source image together in groups when sorted by name in file explorers.
```vbnet
With theMailer
    'Resize
    .ImageSizes = "" ' String Format:  "width, height, suffix | width, height, suffix |". ' Resize only works if sizes defined. 
    .PreventEnlarge = True
    .PreserveAspectRatio = True

    'Crop
    .CropSizes = "" ' String Format:  "width, height | width, height". ' *Must match resize order. Crop only works if sizes defined in Resize and Crop.
    .CropPosition = "Center-Middle"

     'Watermark
     .WatermarkMask = "" ' Add path to a Watermark mask to enable watermarks.
     .WatermarkPadding = 10
     .WatermarkOpacity = 50
     .WatermarkSizes = "128, 128"
     .WatermarkAlign = "Center-Middle"

     'Captions
     .CaptionText = "" Add text to enable captions.
     .CaptionFont = "Ariel"
     .CaptionFontSizes = "16"
     .CaptionFontColor = "Black"
     .CaptionFontStyle =  "Regular" ' Valid values are: "Regular", "Bold", "Italic", "Underline", and "Strikeout".
     .CaptionOpacity = 100
     .CaptionPadding = 10
     .CaptionAlign = "Center-Middle" 
End With
```
When using the upload utility configure any imaging properties you need and call ProcessUploads().  
```vbnet
With theMailer
' Freestanding File Upload & Imaging
ProcessUploads() ' Constructs file upload and imaging without an email message.
End With
```

### Construct `RazorSmartMailer` Instance
Using RazorSmartMailer is very straightforward. True there are many properties the user can define, but only a few the user must define. Most have default values you're unlikely to change. 

To send an HTML templated email, you'll need 
```vbnet
Dim theMailer As New RazorSmartMailer
With theMailer
    .MailTemplatePath = "~/mailtemplates/mailtemplate.vbhtml"
    .SuccessRedirect = "~/thanks.vbhtml" 

    'RazorSmartMailer sendmail properties
    .EmailFrom = "web@razorsmartmailer.com" ' 
    .EmailRecipient = "user@razorsmartmailer.com" 
    .EMailSubject = "Your Email Subject"
    SendWebMail() 
End With
```
We suggest you start setting  only the properties required then tweak the defaults until the desired effect is produced.
### Options
Crop, Watermark, and Caption methods all use the same 9 point location scheme. It is comprised of 3 horizontal and 3 vertical positions resulting in 9 permutations.

Horizontal positions: Left Center Right.

Vertical positions: Top Middle Bottom.

Combine the horizontal and vertical choices in any of the 9 possible combinations to select position.
### Notes

TBD

## Installation
**NuGet**: TBD

## Contributors

* [James Threadgill](https://github.com/wwwebconcepts)

## License

RazorSmartMailer is available under the MIT license. See the [LICENSE](https://github.com/wwwebconcepts/RazorSmartMailer/blob/master/License) file for more info.
