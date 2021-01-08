$error.clear()
$errorpath = "$env:USERPROFILE\Documents\signature-new_error.txt"

#----------------# Domain/Network Dependencies #----------------# 
Write-Host "Checking network dependencies"  -ForegroundColor Yellow
$DC = Get-ADDomainController

if ($? -eq $true){
Write-host "Domain Controller Information: 
Domain Controller: $($DC.Name) 
Domain Controller IP Address: $($DC.IPv4Address)"  -ForegroundColor Green
}else{
Write-Host "System not configured for Active Directory. Exiting Script" -ForegroundColor Red
}

Write-Host "Checking that Domain Controller is reachable"  -ForegroundColor Green
    Test-Connection $DC.IPv4Address -Quiet | Out-Null 
        if ($? -eq $true){
        Write-Host "Domain Controller can be reached, continuing with script" -ForegroundColor Green
        }else{
        Write-Host "Domain Controller cannot be reached. Exiting Script"  -ForegroundColor Red
    }

#----------------# Variable Dependencies #----------------# 

Write-Host "Preparing backend dependencies" -ForegroundColor Yellow
$userid = Get-ADUser -Identity $env:USERNAME -Properties *

#Template File Details
$templateFilePathENV = "$env:USERPROFILE\Downloads\VentanaSignatureCreation\SigFile\template_v3.htm"
$templateFileName = 'template_v3.htm'
$tempSaveLocation = "$env:USERPROFILE"
$sigPhotos = "$env:USERPROFILE\Downloads\VentanaSignatureCreation\SigFile\template_v3_files"

#Outlook Signature Location
$sigFilePath = "$env:USERPROFILE\AppData\Roaming\Microsoft\Signatures\Ventana-New.htm"
$sigFilePathRTF = "$env:USERPROFILE\AppData\Roaming\Microsoft\Signatures\Ventana-New.rtf"
$sigFilePathTXT = "$env:USERPROFILE\AppData\Roaming\Microsoft\Signatures\Ventana-New.txt"
$sigPath = "$env:USERPROFILE\AppData\Roaming\Microsoft\Signatures"
$sigName = "Ventana-New"

 
Write-Host "Backend dependencies completed"  -ForegroundColor Green

#----------------# Outlook Dependencies #----------------# 

Write-Host "Preparing Outlook and system dependencies"  -ForegroundColor Yellow
$UserAccount = Get-ItemProperty -Path HKCU:\Software\Microsoft\Office\Outlook\Settings -Name Accounts | Select -ExpandProperty Accounts
$UserAccountId = (ConvertFrom-Json $UserAccount).UserUpn[0]

#----------------# Find Outlook Profiles in registry #----------------# 
# Find Outlook Profiles in registry
$CommonSettings = $False
$Profiles = (Get-ChildItem HKCU:\Software\Microsoft\Office\16.0\Outlook\Profiles).PSChildName
   Write-Host "Beginning search for Outlook signature registry path" -ForegroundColor Cyan
for ($i = 0; $i -lt 100; $i++){  
    $OutLookProfilePath = "HKCU:\Software\Microsoft\Office\16.0\Outlook\Profiles\" + $Profiles.Trim() + "\9375CFF0413111d3B88A00104B2A6676\0000000$i"
    $OutlookProfile = Get-ItemProperty -Path $OutLookProfilePath -ErrorAction Ignore
    if ($OutlookProfile."Account Name" -eq $UserId.UserPrincipalName){
    Write-Host "UPN found under path $($OutLookProfilePath)" -ForegroundColor Green
    break
}else{
    }
   }
   

#----------------# SigArray Definition CSV #----------------# 
Write-Host "Importing in important variables"  -ForegroundColor Yellow
$SigArray = Import-Csv $env:USERPROFILE\Downloads\VentanaSignatureCreation\SigFile\SigArray.csv

if ($? -eq $true){
Write-Host "Variables import completed"  -ForegroundColor Green
}else{
Write-Host "Variables import not completed. Exiting Script"  -ForegroundColor Red
exit
}
 
#----------------# Signature Cleanup / Old Signature Removal #----------------# 
write-host "Initiating signature cleanup..."  -ForegroundColor Yellow

$i = Test-Path $sigFilePath

if ($i -eq $true){
    Write-Host "Removing old Ventana Signature Files.."  -ForegroundColor Yellow
    Remove-Item $sigFilePath -ErrorAction SilentlyContinue -Force -Confirm:$false
    Remove-Item $sigFilePathRTF -ErrorAction SilentlyContinue -Force -Confirm:$false
    Remove-Item $sigFilePathTXT -ErrorAction SilentlyContinue -Force -Confirm:$false
    Write-Host "Old signature content successfully removed!"  -ForegroundColor Green
}else{
    write-host "No existing Ventana Signatures..."  -ForegroundColor Red
}

#----------------# Signature Base File Creation #----------------#     
Write-Host "Initiating signature creation.."  -ForegroundColor Yellow
New-Item -path $sigFilePath -ItemType file | out-null
Write-Host "Empty signature Created!"  -ForegroundColor Green

#----------------# Signature Media Content / Photos Content #----------------# 
$i = Test-Path $sigPath\Ventana-New_files
if ( $i -eq $true){
    Write-Host "Old signature media exists. Refreshing content"  -ForegroundColor Yellow
    Remove-Item  -Recurse -Force -Confirm:$false $sigPath\Ventana-New_files
    write-host "Copying signature media to destination directory"   -ForegroundColor Yellow
    New-Item -path $sigPath\Ventana-New_files -ItemType Directory | out-null
    Copy-Item -Force -path $sigPhotos\* -Destination $sigPath\Ventana-New_files\
}else{
    Write-Host "Signature media does not exist. Continuing with media migration"  -ForegroundColor Yellow 
    write-host "Copying signature media to destination directory"  -ForegroundColor Yellow
    New-Item -path $sigPath\Ventana-New_files -ItemType Directory | out-null
    Copy-Item -recurse -Force -path $sigPhotos\* -Destination $sigPath\Ventana-New_files\
}


#----------------# Signature Creation #----------------# 

Write-Host "Replacing empty signature with user data.."  -ForegroundColor Yellow

(get-content -path $templateFilePathENV) -replace $SigArray.SignatureTemplateValue[0], $userid.GivenName | Set-Content -Path $sigFilePath
if ($? -eq $true){
}else{
    write-host "Something went wrong.. Exiting"  -ForegroundColor Red
    exit
}

(get-content -path $sigFilePath) -replace $SigArray.SignatureTemplateValue[1], $userid.Surname | Set-Content -Path $sigFilePath
if ($? -eq $true){
}else{
    write-host "Something went wrong.. Exiting"  -ForegroundColor Red
    exit
}

(get-content -path $sigFilePath) -replace $SigArray.SignatureTemplateValue[2], $userid.Description | Set-Content -Path $sigFilePath 
if ($? -eq $true){
}else{
    write-host "Something went wrong.. Exiting"  -ForegroundColor Red
    exit
}

(get-content -path $sigFilePath) -replace $SigArray.SignatureTemplateValue[3], $userid.MobilePhone | Set-Content -Path $sigFilePath
if ($? -eq $true){
}else{
    write-host "Something went wrong.. Exiting"  -ForegroundColor Red
    exit
}

(get-content -path $sigFilePath) -replace $SigArray.SignatureTemplateValue[4], '604.291.9000' | Set-Content -Path $sigFilePath 
if ($? -eq $true){
}else{
    write-host "Something went wrong.. Exiting"  -ForegroundColor Red
    exit
}

(get-content -path $sigFilePath) -replace $SigArray.SignatureTemplateValue[5], $userid.EmailAddress.ToLower() | Set-Content -Path $sigFilePath 
if ($? -eq $true){
}else{
    write-host "Something went wrong.. Exiting"  -ForegroundColor Red
    exit
}
Write-Host "Universalizing signature content!!"  -ForegroundColor Cyan

(get-content -path $sigFilePath) -replace 'template_v3_files', 'Ventana-New_files' | Set-Content -Path $sigFilePath 
write-host "Signature successfully configured!"  -ForegroundColor Green
if ($? -eq $true){
}else{
    write-host "Something went wrong.. Exiting"  -ForegroundColor Red
    exit
}

#----------------#  HTML Signature Conversion #----------------# 

Write-Host "Starting final signature file creation"  -ForegroundColor Yellow

$wrd = new-object -com word.application 
$wrd.visible = $false 
$doc = $wrd.documents.open("$env:USERPROFILE\AppData\Roaming\Microsoft\Signatures\Ventana-New.htm")
$opt = 6
$name = $("$env:USERPROFILE\AppData\Roaming\Microsoft\Signatures\Ventana-New.rtf")
$wrd.ActiveDocument.Saveas([ref]$name,[ref]$opt)
Write-Host "RTF file created"  -ForegroundColor Green
$opt = 6
$name = $("$env:USERPROFILE\AppData\Roaming\Microsoft\Signatures\Ventana-New.txt")
$wrd.ActiveDocument.Saveas([ref]$name,[ref]$opt)
Write-Host "TXT file created"  -ForegroundColor Green
$wrd.Quit()

Write-Host "Applying signatures to Outlook profile"  -ForegroundColor Yellow
Get-Item -Path $OutlookProfilePath | New-Itemproperty -Name "New Signature"  -value $SigName -Propertytype string -Force
if ($? -eq $true){
    Write-Host "Signatures for new emails applied to Outlook profile successfully"  -ForegroundColor Green
 }else{
    write-host "Something went wrong.. Exiting"  -ForegroundColor Red
    exit
    }
    
if ($error.count -gt '0'){
    $error > $errorpath
    Write-Host "Errors found in script. Added to $($errorpath)" -ForegroundColor Red
    }else{
    Write-Host "No errors found in New Signature configuration" -ForegroundColor Green
    write-host "Done" 
    }