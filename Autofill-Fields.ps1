#requires -version 2
<#
.SYNOPSIS
  AutoFill fields in a Microsoft Word document

.DESCRIPTION
  This script will get the active Word document, read the Content Controls and will set the value according to user information stored in Azure Active Directory.

.PARAMETER <Parameter_Name>
    <Brief description of parameter input required. Repeat this attribute if required>

.INPUTS
  None

.OUTPUTS
  None

.NOTES
  Version:          1.1
  Template version: 1.3
  Author:           Axel Timmermans
  Creation Date:    2019
  Purpose/Change:   Updated script formatting - Initial script development 
  
.EXAMPLE
  Use readme.md to deploy, embed in Word Template with macro's

#>

#region Parameters-----------------------------------------------------------------------------------------
<#
Param (
    #Parameters go here
)
#>
#endregion-------------------------------------------------------------------------------------------------

#region Initializations------------------------------------------------------------------------------------
#endregion-------------------------------------------------------------------------------------------------

#region Declarations---------------------------------------------------------------------------------------
#endregion-------------------------------------------------------------------------------------------------

#region Functions------------------------------------------------------------------------------------------

Function Get-UserPrincipalName {
    #Version 1.1, catch has been modified
    try{
        $objUser = New-Object System.Security.Principal.NTAccount($Env:USERNAME)
        $strSID = ($objUser.Translate([System.Security.Principal.SecurityIdentifier])).Value
        $basePath = "HKLM:\SOFTWARE\Microsoft\IdentityStore\Cache\$strSID\IdentityCache\$strSID"
        if((test-path $basePath) -eq $False){
            $userId = $Null
        }
        $userId = (Get-ItemProperty -Path $basePath -Name UserName).UserName
        Return $userId
    }catch{
        Write-Error "Failed to auto detect user principal name"
        Read-Host -Prompt "Press any key to exit"
        Exit
    }
}

#endregion-------------------------------------------------------------------------------------------------


#region Execution------------------------------------------------------------------------------------------

#Get AAD module
    $PackageProvider = Get-PackageProvider -Name NuGet -ListAvailable -ErrorAction SilentlyContinue | Where-Object {$_.Version -ge 2.8}
    $AADModule = Get-InstalledModule -Name AzureAd -MinimumVersion 2.0 -ErrorAction SilentlyContinue
    if(!$PackageProvider){
        Install-PackageProvider -Name NuGet -Scope CurrentUser -Force -Confirm:$false
    }
    if(!$AADModule){
        Install-Module -Name AzureAd -Scope CurrentUser -Force -Confirm:$false
    }
    
#Connect with and get values from AAD. Disconnect after done
    $UPN = Get-UserPrincipalName
    Connect-AzureAD -AccountId $UPN
    $UserInfo = Get-AzureAdUser -ObjectId $UPN
    
    #Contact information - Physical
    $StreetAddress = $UserInfo.StreetAddress
    $City = $UserInfo.City
    $Country = $UserInfo.Country
    $State = $UserInfo.State
    $PostalCode = $UserInfo.PostalCode
        
    #Contact information - Digital
    $Email = $UserInfo.Mail
    $MobilePhone = $UserInfo.Mobile
    $OfficePhone = $UserInfo.TelephoneNumber

    #Company information
    $CompanyName = $UserInfo.CompanyName
    $Department = $UserInfo.Department
    $JobTitle = $UserInfo.JobTitle

    #Names
    $DisplayName = $UserInfo.DisplayName
    $FirstName = $UserInfo.GivenName
    $LastName = $UserInfo.Surname

    Disconnect-AzureAD

#Get the active Word document/template and its sections
    $Word = [Runtime.Interopservices.Marshal]::GetActiveObject('Word.Application')
    $ActiveDoc = $Word.ActiveDocument    
    $StoryRanges = $ActiveDoc.StoryRanges
    Remove-Variable CCs -Force -ErrorAction SilentlyContinue
    foreach($StoryRange in $StoryRanges){
        $CCs += $StoryRange.ContentControls
    }

#loop throug content controls
    foreach($CC in $CCs){
        Switch($CC.Tag){
            #Contact information - Physical
            'StreetAddress'{
                $Item = $word.ActiveDocument.ContentControls.Item($CC.ID)
                $Item.Range.Text = $StreetAddress
            }
            'City'{
                $Item = $word.ActiveDocument.ContentControls.Item($CC.ID)
                $Item.Range.Text = $City
            }
            'Country'{
                $Item = $word.ActiveDocument.ContentControls.Item($CC.ID)
                $Item.Range.Text = $Country
            }
            'State'{
                $Item = $word.ActiveDocument.ContentControls.Item($CC.ID)
                $Item.Range.Text = $State
            }
            'PostalCode'{
                $Item = $word.ActiveDocument.ContentControls.Item($CC.ID)
                $Item.Range.Text = $PostalCode
            }

            #Contact information - Digital
            'Email'{
                $Item = $word.ActiveDocument.ContentControls.Item($CC.ID)
                $Item.Range.Text = $Email
            }
            'MobilePhone'{
                $Item = $word.ActiveDocument.ContentControls.Item($CC.ID)
                $Item.Range.Text = $MobilePhone
            }
            'OfficePhone'{
                $Item = $word.ActiveDocument.ContentControls.Item($CC.ID)
                $Item.Range.Text = $OfficePhone
            }

            #Company information
            'CompanyName'{
                $Item = $word.ActiveDocument.ContentControls.Item($CC.ID)
                $Item.Range.Text = $CompanyName
            }
            'Department'{
                $Item = $word.ActiveDocument.ContentControls.Item($CC.ID)
                $Item.Range.Text = $Department
            }
            'JobTitle'{
                $Item = $word.ActiveDocument.ContentControls.Item($CC.ID)
                $Item.Range.Text = $JobTitle
            }

            #Names
            'DisplayName'{
                $Item = $word.ActiveDocument.ContentControls.Item($CC.ID)
                $Item.Range.Text = $DisplayName
            }
            'FirstName'{
                $Item = $word.ActiveDocument.ContentControls.Item($CC.ID)
                $Item.Range.Text = $FirstName
            }
            'LastName'{
                $Item = $word.ActiveDocument.ContentControls.Item($CC.ID)
                $Item.Range.Text = $LastName
            }
        }
    }

#endregion-------------------------------------------------------------------------------------------------
