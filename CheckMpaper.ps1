<#

CheckMpaper.ps1

    2019-05-08 Initial Creation

#>

if (!($env:PSModulePath -match 'C:\\PowerShell\\_Modules')) {
    $env:PSModulePath = $env:PSModulePath + ';C:\PowerShell\_Modules\'
}

Get-Module -ListAvailable WorldJournal.* | Remove-Module -Force
Get-Module -ListAvailable WorldJournal.* | Import-Module -Force

$scriptPath = $MyInvocation.MyCommand.Path
$scriptName = (($MyInvocation.MyCommand) -Replace ".ps1")
$hasError   = $false

$newlog     = New-Log -Path $scriptPath -LogFormat yyyyMMdd-HHmmss
$log        = $newlog.FullName
$logPath    = $newlog.Directory

$mailFrom   = (Get-WJEmail -Name noreply).MailAddress
$mailPass   = (Get-WJEmail -Name noreply).Password
$mailTo     = (Get-WJEmail -Name lyu).MailAddress
$mailSbj    = $scriptName
$mailMsg    = ""

$localTemp = "C:\temp\" + $scriptName + "\"
if (!(Test-Path($localTemp))) {New-Item $localTemp -Type Directory | Out-Null}

Write-Log -Verb "LOG START" -Noun $log -Path $log -Type Long -Status Normal
Write-Line -Length 50 -Path $log

###################################################################################




$workDate = (Get-Date).AddDays(0)
$workDay  = ($workDate).DayOfWeek.value__ # 0, 1, 2, 3, 4, 5, 6
if($workDay -eq 0){ $pubcode = @('AT','BO','CH','DC','NJ','NY','NW')
}else{ $pubcode = @('AT','BO','CH','DC','NJ','NY') }
$ftp = Get-WJFTP -Name Oohla

Write-Log -Verb "workDate" -Noun $workDate.ToString("yyyyMMdd") -Path $log -Type Short -Status Normal
Write-Log -Verb "workDay" -Noun $workDay -Path $log -Type Short -Status Normal
Write-Log -Verb "pubcode" -Noun ([string]$pubcode) -Path $log -Type Short -Status Normal
Write-Log -Verb "ftp" -Noun $ftp.Path -Path $log -Type Short -Status Normal

$pubcode | ForEach-Object{

    $pub = $_
    $remotePath = $ftp.Path + $pub + '/'
    $wl = WebRequest-ListDirectory -Username $ftp.User -Password $ftp.Pass -RemoteFolderPath $remotePath

    Write-Log -Verb "CHECK" -Noun $remotePath -Path $log -Type Long -Status Normal
    Write-Log -Verb $wl.Verb -Noun $wl.List -Path $log -Type Long -Status $wl.Status

    if($wl.Status -eq "Bad"){

        $mailMsg = $mailMsg + (Write-Log -Verb $wl.Verb -Noun $wl.Noun -Path $log -Type Long -Status $wl.Status -Output String) + "`n"
        $hasError = $true

    }else{

        if(($wl.List).Count -gt 0){

            $mailMsg = $mailMsg + (Write-Log -Verb $remotePath -Noun ((($wl.List).Count).ToString()+" files") -Path $log -Type Long -Status Good -Output String) + "`n"

        }else{

            $mailMsg = $mailMsg + (Write-Log -Verb $remotePath -Noun ((($wl.List).Count).ToString()+" files") -Path $log -Type Long -Status Bad -Output String) + "`n"
            $hasError = $true

        }

    }

}









###################################################################################

Write-Line -Length 50 -Path $log

# Delete temp folder

Write-Log -Verb "REMOVE" -Noun $localTemp -Path $log -Type Long -Status Normal
try{
    $temp = $localTemp
    Remove-Item $localTemp -Recurse -Force -ErrorAction Stop
    Write-Log -Verb "REMOVE" -Noun $temp -Path $log -Type Long -Status Good
}catch{
    $mailMsg = $mailMsg + (Write-Log -Verb "REMOVE" -Noun $temp -Path $log -Type Long -Status Bad -Output String) + "`n"
    $mailMsg = $mailMsg + (Write-Log -Verb "Exception" -Noun $_.Exception.Message -Path $log -Type Short -Status Bad -Output String) + "`n"
}

Write-Line -Length 50 -Path $log
Write-Log -Verb "LOG END" -Noun $log -Path $log -Type Long -Status Normal
if($hasError){ $mailSbj = "ERROR " + $mailSbj }

$emailParam = @{
    From    = $mailFrom
    Pass    = $mailPass
    To      = $mailTo
    Subject = $mailSbj
    Body    = $mailMsg
    ScriptPath = $scriptPath
    Attachment = $log
}

mailv2 @emailParam