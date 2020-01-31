Add-type -assembly "Microsoft.Office.Interop.Outlook" | out-null

# variables
$NewPSTFolder = "F:\PST"
if (!(Test-Path -Path $NewPSTFolder)) {
    New-Item -Path $NewPSTFolder -ItemType Directory
}

# log paths
$LogPathMountedPST = "$NewPSTFolder\log_MountedPST.txt"
$LogPathCopiedPST = "$NewPSTFolder\log_CopiedPST.txt"
$LogRemountPST = "$NewPSTFolder\log_RemountPst.txt"
$LogPath = "$NewPSTFolder\log.log"

# find PST and unmount it
function unmount-pst {
    $Outlook = new-object -com outlook.application 
    $Namespace = $Outlook.getNamespace("MAPI")
    $MountedPST = $Outlook.Session.Stores | Where-Object -Property FilePath -Like "*.pst"

    write-host "Unmounted PST ran on $(Get-Date)" 
    if ($MountedPST -ne $null) {
        # log action
        $MountedPST | Select-Object -Property FilePath | Out-File -FilePath $LogPathMountedPST -Force -Append
        
        $MountedPST_StringArray = @()
        $MountedPST | Select-Object -Property FilePath | ForEach-Object {
            $MountedPST_StringArray += $_.FilePath
        }
    }
    else {
        write-host "No PST files mounted for this user" 
    }

    foreach ($PST in $MountedPST){
        # Unmount PST
        $PSTRoot = $PST.GetRootFolder()
        $PSTFolder = $namespace.Folders.Item($PSTRoot.Name)
        $namespace.GetType().InvokeMember('RemoveStore',[System.Reflection.BindingFlags]::InvokeMethod,$null,$namespace,($PSTFolder)) | Out-Null
    }  

    # Close Outlook object    
    Stop-Process -name "outlook" -Verbose -ErrorAction SilentlyContinue
    
    # Close Windows Search Indexer process
    # ????????????????????????????????????

return $MountedPST_StringArray}


# copy PST to new location
function copy-pst {
    [CmdletBinding()] # enable all the def params
    Param(
        [Parameter(Mandatory=$true,
            ValueFromPipeline=$true,
            ValueFromPipelineByPropertyName=$true,
            Position=0)]
        [String[]]$PSTs
    )

    write-host "Copying PST ran on $(Get-Date)"

    if (!($PSTs.count -gt 0)) {
        write-host "0 PST were passed as arguements to move"
        return
    }

    Write-Host "Closing Outlook and Skype if not closed in previous function"
    # (get-process outlook).CloseMainWindow() to be tested --
    Stop-Process -name "outlook" -Verbose -ErrorAction SilentlyContinue
    Start-Sleep -Seconds 2

    [String[]]$NewPSTPath = @()
    foreach ($PST in $PSTs){
        #Move PST
        $NewPath = "$NewPSTLocation\$(Split-Path -Path $PST -Leaf)"
        $NewPSTPath += $NewPath
        if ($PST -eq $NewPath) {
            write-host "Source: $PST and Destionation: $NewPath match. Not copying."            
        }
        else {
            write-host "Copying from $PST to $NewPath"
            Copy-Item -Path $PST -Destination $NewPath -Verbose -Force
        }
    }
    
    return $NewPSTPath
}


# mount PST to match new location
function mount-pst {
    [CmdletBinding()] #Enable all the default paramters, including -Verbose
    Param(
        [Parameter(Mandatory=$true,
            ValueFromPipeline=$true,
            ValueFromPipelineByPropertyName=$true,
            Position=0)]
        [String[]]$MountPSTs
    )

    $Outlook = new-object -com outlook.application 
    $Namespace = $Outlook.getNamespace("MAPI")

    foreach ($PST in $MountPSTs){
        #Remount PST
        write-host "Mounting PST: $Pst"
        $namespace.AddStore($PST)
    }  
    
}


# runbook
Start-Transcript -Path $LogPath -Append
    $Unmount = unmount-pst
    $Unmount | Out-File -FilePath $LogPathMountedPST -Force

    $CopyPath = copy-pst -PSTs $Unmount
    $CopyPath | Out-File -FilePath $LogPathCopiedPST -Force 

    mount-pst -MountPSTs $CopyPath

    write-host "Stopping Outlook and returning session to user"
    Stop-Process -name "outlook" -Verbose -ErrorAction SilentlyContinue
Stop-Transcript