#Author  : Shahul Hameed Abdulkareem
###########################################################
# Read me
#1. Set the location of the input folder edt the first line of the source code below
#01-06-2025 - Baseline
#
Set-Location -Path  C:\Test
$date = get-date
$date =$date.Tostring("yyyyMMdd") 

function Get-InstalledApps {
    param (
        [Parameter(ValueFromPipeline=$true)]
        [string[]]$ComputerName = $env:COMPUTERNAME,
        [string]$NameRegex = ''
    )
    
    foreach ($comp in $ComputerName) {
        $keys = '','\Wow6432Node'
        foreach ($key in $keys) {
            try {
                $reg = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey('LocalMachine', $comp)
                $apps = $reg.OpenSubKey("SOFTWARE$key\Microsoft\Windows\CurrentVersion\Uninstall").GetSubKeyNames()
            } catch {

                continue
            }

            foreach ($app in $apps) {
                $program = $reg.OpenSubKey("SOFTWARE$key\Microsoft\Windows\CurrentVersion\Uninstall\$app")
                $name = $program.GetValue('DisplayName')
                if ($name -and $name -match $NameRegex) {
                    [pscustomobject]@{
                        ComputerName = $comp
                        DisplayName = $name
                        DisplayVersion = $program.GetValue('DisplayVersion')
                        Publisher = $program.GetValue('Publisher')
                        InstallDate = $program.GetValue('InstallDate')
                        UninstallString = $program.GetValue('UninstallString')
                        Bits = $(if ($key -eq '\Wow6432Node') {'64'} else {'32'})
                        Path = $program.name
                    }
                }
            }
        }
    }
}

function Get-ProcessV1
{
Get-CimInstance -ClassName Win32_Process  |Where-Object { 
    (Get-Process -Id $_.ProcessId).TotalProcessorTime -ne $null 
} |Select Name,ProcessId,Path,CommandLine

}


function Get-PortStatus
{


$DB = Get-Content Port.csv | select-object -skip 1 #skip the first header row


$Resultlist = new-object system.collections.arraylist



$newfilename = $env:COMPUTERNAME  +"_PortInfo_" +$date +".csv" 


New-Item   $newfilename -ItemType file -Force


$OldServer =''
foreach ($Data in $DB) {
  $Source, $Port, $Destination = $Data -split ',' -replace '^\s*|\s*$'
  
  write-host "SourceServer is: "$Source
  write-host "Port is: "$Port
  write-host "Destination is: "$Destination
  #If (-NOT ($Source -eq  $OldServer)) {
  write-host "Test"
    # Exit-PSSession -
    # Enter-PSSession -ComputerName $Source 
 # }
    If ( Test-Connection $Destination -Count 1 -Quiet) {
    
        try {       
            $null = New-Object System.Net.Sockets.TCPClient -ArgumentList $Destination,$Port
            $props = @{
                SourceServer = $Source
                Port = $Source
                DestinationServer = $Destination
                PortOpen = 'Yes'
            }# | Format-Table | Out-File -FilePath $newfilename -Append 
            $PortOpen = 'Yes'  
            
        }

        catch {
            $props = @{
                SourceServer = $Source
                Port = $Port
                DestinationServer = $Destination
                PortOpen = 'No'
            } #| Format-Table | Out-File -FilePath  $newfilename -Append 
            $PortOpen = 'No'
        }
    }

    Else {
        
        $props = @{
            SourceServer = $Source
            Port = $Port
            DestinationServer = $Destination
            PortOpen = 'Server did not respond to ping'
        } #| Format-Table | Out-File -FilePath  $newfilename -Append 
        $PortOpen = 'Server did not respond to ping'
    }

    $Output = New-Object -Type PSCustomObject 
    $output | Add-Member -MemberType NoteProperty -Name "Source" -value $Source
    $output | Add-Member -MemberType NoteProperty -Name "Port" -value $Port
    $output | Add-Member -MemberType NoteProperty -Name "Destination" -value $Destination
    $output | Add-Member -MemberType NoteProperty -Name "PortOpen" -value $PortOpen

    $Resultlist.add($Output) | Out-Null
    New-Object PsObject -Property $props
    $OldServer = $Source
  
}


$Resultlist | Export-Csv $newfilename -NoTypeInformation

$Resultlist[0][0][0][0]
}


function Convert-WuaResultCodeToName
{
    param(
        [Parameter(Mandatory=$true)]
        [int] $ResultCode
    )

    $Result = $ResultCode
    switch($ResultCode)
    {
      2 {
        $Result = "Succeeded"
      }
      3 {
        $Result = "Succeeded With Errors"
      }
      4 {
        $Result = "Failed"
      }
    }

    return $Result
}

function Get-WuaHistory
{

  # Get a WUA Session
  $session = (New-Object -ComObject 'Microsoft.Update.Session')

  # Query the latest 1000 History starting with the first recordp     
  $history = $session.QueryHistory("",0,1000) | ForEach-Object {
     $Result = Convert-WuaResultCodeToName -ResultCode $_.ResultCode

     # Make the properties hidden in com properties visible.
     $_ | Add-Member -MemberType NoteProperty -Value $Result -Name Result
     $Product = $_.Categories | Where-Object {$_.Type -eq 'Product'} | Select-Object -First 1 -ExpandProperty Name
     $_ | Add-Member -MemberType NoteProperty -Value $_.UpdateIdentity.UpdateId -Name UpdateId
     $_ | Add-Member -MemberType NoteProperty -Value $_.UpdateIdentity.RevisionNumber -Name RevisionNumber
     $_ | Add-Member -MemberType NoteProperty -Value $Product -Name Product -PassThru

     Write-Output $_
  } 

  #Remove null records and only return the fields we want
  $history | 
      Where-Object {![String]::IsNullOrWhiteSpace($_.title)} | 
          Select-Object Result, Date, Title, SupportUrl, Product, UpdateId, RevisionNumber
}  

function Get-ComInfo
{

$comAdmin = New-Object -comobject COMAdmin.COMAdminCatalog
$apps = $comAdmin.GetCollection("Applications")
$apps.Populate()

$appInstances = $comAdmin.GetCollection("ApplicationInstances")
$appInstances.Populate()

foreach ($appInstance in $appInstances) {
  # identify the associated application
  $appKey = $appInstance.Value('Application')
  $appName = ($apps |Where-Object Key -eq $appKey).Name

  # output along with hosting process id
  [pscustomobject]@{
    ApplicationName = $appName
    ProcessId = $appInstance.Value('ProcessID')
  }
}

}




$filenamev1= $env:COMPUTERNAME  +"_ComputerInfo_"+$date +".txt"
$ServiceInfo= $env:COMPUTERNAME  +"_ServiceInfo_" +$date +".csv"
$WindowsUpdatesInfo= $env:COMPUTERNAME  +"_WindowsUpdatesInfo_" +$date +".csv"
$InstalledAppsInfo= $env:COMPUTERNAME  +"_InstalledAppsInfo_" +$date +".csv"
$ProcessListInfo= $env:COMPUTERNAME  +"_ProcessListInfo_" +$date +".csv"
$ComprocessListInfo= $env:COMPUTERNAME  +"_ComprocessListInfo_" +$date +".csv"

Get-PortStatus 

Get-ComputerInfo  | Out-File -FilePath $filenamev1
Get-NetIPConfiguration | Out-File -FilePath $filenamev1  -Append	
Get-NetIPInterface | Out-File -FilePath $filenamev1  -Append	
Get-Service |  Export-CSV -Path $ServiceInfo
Get-InstalledApps  -ComputerName  $env:COMPUTERNAME  |  Export-CSV -Path $InstalledAppsInfo  -NoTypeInformation
Get-WuaHistory |  Export-CSV -Path $WindowsUpdatesInfo -NoTypeInformation
Get-ProcessV1 | Export-CSV -Path $ProcessListInfo -NoTypeInformation
Get-ComInfo | Export-CSV -Path $ComprocessListInfo -NoTypeInformation