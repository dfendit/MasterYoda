Add-type -assembly "Microsoft.Office.Interop.Outlook" | out-null

# variables
$NewPSTFolder = "F:\PST"
if (!(Test-Path -Path $NewPSTFolder)) {
    New-Item -Path $NewPSTFolder -ItemType Directory
}

# log paths
$LogPathMountedPST = "$NewPSTFolder\log_MountedPST.txt"
#$LogPathCopiedPST = "$NewPSTFolder\log_CopiedPST.txt"
#$LogRemountPST = "$NewPSTFolder\log_RemountPst.txt"
$LogPath = "$NewPSTFolder\log.log"

# find PST and unmount it
function unmount-pst {
    $Outlook = new-object -com outlook.application 
    $Namespace = $Outlook.getNamespace("MAPI")
    $MountedPST = $Outlook.Session.Stores | Where-Object -Property FilePath -Like "*.pst"

    write-host "Unmounted PST ran on $(Get-Date)" 
    if ($MountedPST -ne $null) {
        #debug
        $MountedPST | Select-Object -Property FilePath | Out-File -FilePath $LogPathMountedPST -Force -Append
        
        $MountedPST_StringArray = @()
        $MountedPST | Select-Object -Property FilePath | ForEach-Object {
            $MountedPST_StringArray += $_.FilePath
        }
    }
    else {
        write-host "No PST files mounted" 
    }

    foreach ($PST in $MountedPST){
        #Unmount PST
        $PSTRoot = $PST.GetRootFolder()
        $PSTFolder = $namespace.Folders.Item($PSTRoot.Name)
        $namespace.GetType().InvokeMember('RemoveStore',[System.Reflection.BindingFlags]::InvokeMethod,$null,$namespace,($PSTFolder)) | Out-Null
    }  

    #Close Outlook    
    Stop-Process -name "Outlook"
    Remove-Variable Outlook | Out-Null

    #Close Windows Search Indexer process


return $MountedPST_StringArray}


# copy PST to new location



# mount PST to match new location
function mount-pst {
    [CmdletBinding()] #Enable all the default paramters, including -Verbose
    Param(
        [Parameter(Mandatory=$true,
            ValueFromPipeline=$true,
            ValueFromPipelineByPropertyName=$true,
            Position=0)]
        [String[]]$PSTs
    )

    $Outlook = new-object -com outlook.application 
    $Namespace = $Outlook.getNamespace("MAPI")

    foreach ($PST in $PSTs){
        #Remount PST
        write-host "Mounting PST: $Pst"
        $namespace.AddStore($PST)
    }  
    
}


# runbook
Start-Transcript -Path $LogPath -Append
    $Temp = unmount-pst
    $Temp | Out-File -FilePath $LogPathMountedPST -Force
    
 
Stop-Transcript