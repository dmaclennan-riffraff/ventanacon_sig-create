$ADTools = Get-WindowsCapability -Online -Name Rsat.ActiveDirectory.DS-LDS.Tools~~~~0.0.1.0

if ($ADTools.State -eq 'NotPresent'){
    Write-Host "Active Directory LWT not installed."
    Write-Host "Attempting first install of Active Directory LWT"
    Add-WindowsCapability -Online -Name Rsat.ActiveDirectory.DS-LDS.Tools~~~~0.0.1.0
    $ADTools = Get-WindowsCapability -Online -Name Rsat.ActiveDirectory.DS-LDS.Tools~~~~0.0.1.0
}else{}
        
if ($ADTools.State -eq "Installed"){
    write-host "Active Directory LWT is installed. Continuing with script" 
        }else{
            Write-Host "Active Directory LWT is still not installed. Exiting script" 
        }
            
$ExecPol = Set-ExecutionPolicy RemoteSigned -Force

if ($ExecPol -eq "RemoteSigned"){
    Write-Host "Execution policy successfully set to RemoteSigned"
        }else{
            write-host "Execution policy not set correctly. Cannot continue with scripts"
        }

    
