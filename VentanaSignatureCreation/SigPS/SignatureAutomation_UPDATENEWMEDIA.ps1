$SigArray = Import-Csv $env:USERPROFILE\Downloads\VentanaSignatureCreation\SigFile\SigArray.csv

Write-Host "Please update the following settings to make changes to the New Signature media" 

$SigUpdate = Read-Host "Would you like to make changes to image001? (y/n)"

if ($SigUpdate -eq "y"){
    $image001 = Read-Host "Please provide a new URL for image001"
    (Get-Content $env:USERPROFILE\Downloads\VentanaSignatureCreation\SigFile\SigArray.csv) -replace $SigArray.PhotoUrl[1], $image001 | Set-Content $env:USERPROFILE\Downloads\VentanaSignatureCreation\SigFile\SigArray.csv
    write-host "Image001 has been updated"
    }else{
    write-host "Image001 has not been updated"
 }

 $SigUpdate = Read-Host "Would you like to make changes to image002? (y/n)"
     if ($SigUpdate -eq "y"){
    $image002 = Read-Host "Please provide a new URL for image002"
    (Get-Content $env:USERPROFILE\Downloads\VentanaSignatureCreation\SigFile\SigArray.csv) -replace $SigArray.PhotoUrl[2], $image002 | Set-Content $env:USERPROFILE\Downloads\VentanaSignatureCreation\SigFile\SigArray.csv
    write-host "Image002 has been updated"
    }else{
    write-host "Image002 has not been updated"
}

$SigUpdate = Read-Host "Would you like to make changes to image003? (y/n)"

    if ($SigUpdate -eq "y"){
    $image003 = Read-Host "Please provide a new URL for image003"
    (Get-Content $env:USERPROFILE\Downloads\VentanaSignatureCreation\SigFile\SigArray.csv) -replace $SigArray.PhotoUrl[3], $image003 | Set-Content $env:USERPROFILE\Downloads\VentanaSignatureCreation\SigFile\SigArray.csv
    write-host "Image003 has been updated"
    }else{
    write-host "Image003 has not been updated"
 }


 $SigUpdate = Read-Host "Would you like to make changes to image004? (y/n)"

    if ($SigUpdate -eq "y"){
    $image004 = Read-Host "Please provide a new URL for image004"
    (Get-Content $env:USERPROFILE\Downloads\VentanaSignatureCreation\SigFile\SigArray.csv) -replace $SigArray.PhotoUrl[4], $image004 | Set-Content $env:USERPROFILE\Downloads\VentanaSignatureCreation\SigFile\SigArray.csv
    write-host "Image004 has been updated"
    }else{
    write-host "Image004 has not been updated"
 }

 $SigUpdate = Read-Host "Would you like to make changes to image005?(y/n)"

    if ($SigUpdate -eq "y"){
    $image005 = Read-Host "Please provide a new URL for image005"
    (Get-Content $env:USERPROFILE\Downloads\VentanaSignatureCreation\SigFile\SigArray.csv) -replace $SigArray.PhotoUrl[5], $image005 | Set-Content $env:USERPROFILE\Downloads\VentanaSignatureCreation\SigFile\SigArray.csv
    write-host "Image005 has been updated"
    }else{
    write-host "Image005 has not been updated"
  }

Write-Host "Script completed. Exiting."
    exit