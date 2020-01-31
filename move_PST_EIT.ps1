Add-type -assembly "Microsoft.Office.Interop.Outlook" | out-null

$NewPSTFolder = "F:\PST"

if (!(Test-Path -Path $NewPSTFolder)) {
    New-Item -Path $NewPSTFolder -ItemType Directory
}

$LogPathMountedPST = "$NewPSTFolder\log_MountedPST.txt"
#$MovedPSTSRecord = "$NewPSTFolder\MovedPSTs.txt"
#$RemountedPSTSRecord = "$NewPSTFolder\RemountedPSTs.txt"
$LogPath = "$NewPSTFolder\log.log"

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
    $Service = Get-Service -Name 'WSearch'
        if ($pscmdlet.ShouldProcess($Service)) {
            if (!($Enable.IsPresent)) {
                if ($Service.Status -eq 'Running') {
                    $Service | Stop-Service -Force
                }
                $Service | Set-Service -StartupType Disable
            } else {
                $Service | Set-Service -StartupType Automatic
                $Service | Start-Service
            }
        }

    return $MountedPST_StringArray}

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

    Start-Transcript -Path $LogPath -Append
    $Temp = unmount-pst
    $Temp | Out-File -FilePath $MountedPSTSRecord -Force
    
 
    Stop-Transcript