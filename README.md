# Autofill for Microsoft Office documents
Autofill fields in Microsoft Office documents with (user) data from AzureAD.
Currently, only Word is supported, but adding support for other applications should not be too much work. Feel free to create an issue for this or create a pull request.

## How to use
### For developers
#### Fields in the document
Fill the field tag of the Content Control object. The tag should match one of the following:
* Contact information - Physical
  * StreetAddress
  * City
  * Country
  * State
  * PostalCode
* Contact information – Digital
  * Email
  * MobilePhone
  * OfficePhone
* Company information
  * CompanyName
  * Department
  * JobTitle
* Names
  * DisplayName
  * FirstName
  * LastName
  
#### Device configuration
* Sync the library that contains the PowerShell script for the autofill
* Set trusted locations for the template files (https://github.com/wiseleaf23/microsoft-device-management/tree/master/Microsoft%20Office%20client%20apps)
* Exclude the location that contains the template files via Windows Defender, or sign the template files (signing has not been tested yet)

#### VBA in the document
You should add the following lines of VBA to your document:
```vb
Sub AutoNew()
    strCommand = "Powershell -ExecutionPolicy ByPass -WindowStyle Hidden -File ""%UserProfile%\<AAD name>\<synced library>\Autofill-Fields.ps1"""
    Set objShell = CreateObject("WScript.Shell")
    strErrorCode = objShell.Run(strCommand, 0, True)
    strErrorCode = objShell.Run(strCommand, 0, False)
End Sub
```
Replace `<AAD name>\<synced library>` with the path to the synced SharePoint library. If the location is not a synced SharePoint library, you should replace the entire path. I recommend storing the PowerShell script in the same library as the templates and making this read only. This way, you can easily update the templates and the code.

### For end-users
Please note the following:
* Autofill will only work the first time you open a document from the template
* The first time you open a template with autofill, you could get a warning about the code that is contained in the document. You should mark this document as trusted, but always be careful about which document to mark as trusted!
* The first time it runs will take a longer time because components will be installed, depending on the speed of your computer and your internet it should take about 20 seconds
* After the first run, autofill should take no more than 5 seconds
* You need to have a working internet connection

