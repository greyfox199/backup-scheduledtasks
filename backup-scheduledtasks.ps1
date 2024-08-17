[cmdletbinding()]
param (
    [Parameter (Mandatory = $true)] [String]$ConfigFilePath
)

#if json config file does not exist, abort process
if (-not(Test-Path -Path $ConfigFilePath -PathType Leaf)) {
    throw "json config file specified at $($ConfigFilePath) does not exist, aborting process"
}
  
#if config file configured is not json format, abort process.
try {
    $PowerShellObject=Get-Content -Path $ConfigFilePath | ConvertFrom-Json
} catch {
    throw "Config file of $($ConfigFilePath) is not a valid json file, aborting process"
}

#if scheduled task path option does not exist in json, abort process
if ($PowerShellObject.Required.ScheduledTaskPath) {
    $ScheduledTaskPath = $PowerShellObject.Required.ScheduledTaskPath
} else {
    throw "ScheduledTaskPath does not exist in json config file, aborting process"
}

#if scheduled task path option does not exist in json, abort process
if ($PowerShellObject.Required.LocalBackupDirectory) {
    $LocalBackupDirectory = $PowerShellObject.Required.LocalBackupDirectory
} else {
    throw "LocalBackupDirectory does not exist in json config file, aborting process"
}

#if errorMailSender optoin does not exist in json, abort process
if ($PowerShellObject.Required.errorMailSender) {
    $errorMailSender = $PowerShellObject.Required.errorMailSender
} else {
    throw "errorMailSender does not exist in json config file, aborting process"
}

#if errorMailRecipients option does not exist in json, abort process
if ($PowerShellObject.Required.errorMailRecipients) {
    $errorMailRecipients = $PowerShellObject.Required.errorMailRecipients
} else {
    throw "errorMailRecipients does not exist in json config file, aborting process"
}

#if errorMailTenantID option does not exist in json, abort process
if ($PowerShellObject.Required.errorMailTenantID) {
    $errorMailTenantID = $PowerShellObject.Required.errorMailTenantID
} else {
    throw "errorMailTenantID does not exist in json config file, aborting process"
}

#if errorMailAppID option does not exist in json, abort process
if ($PowerShellObject.Required.errorMailAppID) {
    $errorMailAppID = $PowerShellObject.Required.errorMailAppID
} else {
    throw "errorMailAppID does not exist in json config file, aborting process"
}

#if errorMailSubjectPrefix option does not exist in json, abort process
if ($PowerShellObject.Required.errorMailSubjectPrefix) {
    $errorMailSubjectPrefix = $PowerShellObject.Required.errorMailSubjectPrefix
} else {
    throw "errorMailSubjectPrefix does not exist in json config file, aborting process"
}

#if errorMailPasswordFile option does not exist in json, abort process
if ($PowerShellObject.Required.errorMailPasswordFile) {
    $errorMailPasswordFile = $PowerShellObject.Required.errorMailPasswordFile
} else {
    throw "errorMailPasswordFile does not exist in json config file, aborting process"
}

#set up variables
[string] $strServerName = $env:computername
[bool] $blnWriteToLog = $false
[int] $intErrorCount = 0
$arrStrErrors = @()

#clear all errors before starting
$error.Clear()

[uint16] $intDaysToKeepLogFiles = 0
[string] $strServerName = $env:computername

#if path to log directory exists, set logging to true and setup log file
if (Test-Path -Path $PowerShellObject.Optional.logsDirectory -PathType Container) {
    $blnWriteToLog = $true
    [string] $strTimeStamp = $(get-date -f yyyy-MM-dd-hh_mm_ss)
    [string] $strDetailLogFilePath = $PowerShellObject.Optional.logsDirectory + "\backup-scheduledtasks-detail-" + $strTimeStamp + ".log"
    $objDetailLogFile = [System.IO.StreamWriter] $strDetailLogFilePath
}

#if days to keep log files directive exists in config file, set configured days to keep log files
if ($PowerShellObject.Optional.daysToKeepLogFiles) {
    try {
        $intDaysToKeepLogFiles = $PowerShellObject.Optional.daysToKeepLogFiles
        Out-GVLogFile -LogFileObject $objDetailLogFile -WriteToLog $true -LogString "$(get-date) Info: Using $($PowerShellObject.Optional.daysToKeepLogFiles) value specified in config file for log retention" -LogType "Info" -DisplayInConsole $false
    } catch {
        Out-GVLogFile -LogFileObject $objDetailLogFile -WriteToLog $true -LogString "$(get-date) Warning: $($PowerShellObject.Optional.daysToKeepLogFiles) value specified in config file is not valid, defaulting to unlimited log retention" -LogType "Warning"
    }
}

Out-GVLogFile -LogFileObject $objDetailLogFile -WriteToLog $true -LogString "$(get-date) Info: Beginning process to backup all scheduled tasks in the $($ScheduledTaskPath) path to $($strLocalDirectory)" -LogType "Info"

Out-GVLogFile -LogFileObject $objDetailLogFile -WriteToLog $true -LogString "$(get-date) Info: Getting all tasks in task path" -LogType "Info"
try {
    $objScheduledTasks = Get-ScheduledTask -TaskPath $ScheduledTaskPath -ErrorAction Stop
    Out-GVLogFile -LogFileObject $objDetailLogFile -WriteToLog $true -LogString "$(get-date) Info: Successfully obtained all tasks in taskpath.  Obtained $($objScheduledTasks.count) tasks" -LogType "Info"
    Out-GVLogFile -LogFileObject $objDetailLogFile -WriteToLog $true -LogString "$(get-date) Info: Looping through each task and exporting to file with taskname in file name" -LogType "Info"
	foreach ($objScheduledTask in $objScheduledTasks) {
        Out-GVLogFile -LogFileObject $objDetailLogFile -WriteToLog $true -LogString "$(get-date) Info: Exporting $($objScheduledTask.TaskName)" -LogType "Info"
		Export-ScheduledTask -TaskName $objScheduledTask.TaskName -TaskPath $objScheduledTask.TaskPath | Out-File "$($LocalBackupDirectory)\$($objScheduledTask.TaskName).xml"
	}
    Out-GVLogFile -LogFileObject $objDetailLogFile -WriteToLog $true -LogString "$(get-date) Info: Successfully exported all tasks in taskpath" -LogType "Info"
} catch {
    $ErrorMessage = $_.Exception.Message
	$line = $_.InvocationInfo.ScriptLineNumber
	$arrStrErrors += "Failed to export all scheduled tasks in the $($ScheduledTaskPath) path to $($strLocalDirectory) at $($line) with the following error: $ErrorMessage"
    Out-GVLogFile -LogFileObject $objDetailLogFile -WriteToLog $true -LogString "$(get-date) Error: Error: Failed to export all scheduled tasks in the $($ScheduledTaskPath) path to $($strLocalDirectory) at $($line) with the following error: $ErrorMessage" -LogType "Error"
}

#log retention
if ($intDaysToKeepLogFiles -gt 0) {
    try {
        Out-GVLogFile -LogFileObject $objDetailLogFile -WriteToLog $blnWriteToLog -LogString "$(get-date) Info: Purging log files older than $($intDaysToKeepLogFiles) days from $($PowerShellObject.Optional.logsDirectory)" -LogType "Info"
        $CurrentDate = Get-Date
        $DatetoDelete = $CurrentDate.AddDays("-$($intDaysToKeepLogFiles)")
        Get-ChildItem "$($PowerShellObject.Optional.logsDirectory)" | Where-Object { $_.LastWriteTime -lt $DatetoDelete } | Remove-Item -Force
    } catch {
        $ErrorMessage = $_.Exception.Message
        $line = $_.InvocationInfo.ScriptLineNumber
        $arrStrErrors += "Failed to purge log files older than $($intDaysToKeepLogFiles) days from $($PowerShellObject.Optional.logsDirectory) with the following error: $ErrorMessage"
        Out-GVLogFile -LogFileObject $objDetailLogFile -WriteToLog $blnWriteToLog -LogString "$(get-date) Error: Failed to purge log files older than $($intDaysToKeepLogFiles) days from $($PowerShellObject.Optional.logsDirectory) with the following error: $ErrorMessage" -LogType "Error"
    }
}

[int] $intErrorCount = $arrStrErrors.Count

if ($intErrorCount -gt 0) {
    Out-GVLogFile -LogFileObject $objDetailLogFile -WriteToLog $blnWriteToLog -LogString "$(get-date) Info: Encountered $intErrorCount errors, sending error report email" -LogType "Error"
    #loop through all errors and add them to email body
    foreach ($strErrorElement in $arrStrErrors) {
        $intErrorCounter = $intErrorCounter + 1
        $strEmailBody = $strEmailBody + $intErrorCounter.toString() + ") " + $strErrorElement + "<br>"
    }
    $strEmailBody = $strEmailBody + "<br>Please see $strDetailLogFilePath on $strServerName for more details"

    Out-GVLogFile -LogFileObject $objDetailLogFile -WriteToLog $blnWriteToLog -LogString "$(get-date) Info: Sending email error report via $($errorMailAppID) app on $($errorMailTenantID) tenant from $($errorMailSender) to $($errorMailRecipients) as specified in config file" -LogType "Info"
    $errorEmailPasswordSecure = Get-Content $errorMailPasswordFile | ConvertTo-SecureString
    $BSTR = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($errorEmailPasswordSecure)
    $errorEmailPassword = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto($BSTR)
    [Runtime.InteropServices.Marshal]::ZeroFreeBSTR($BSTR)

    Send-GVMailMessage -sender $errorMailSender -TenantID $errorMailTenantID -AppID $errorMailAppID -subject "$($errorMailSubjectPrefix): Encountered $($intErrorCount) errors during process" -body $strEmailBody -ContentType "HTML" -Recipient $errorMailRecipients -ClientSecret $errorEmailPassword
}

Out-GVLogFile -LogFileObject $objDetailLogFile -WriteToLog $blnWriteToLog -LogString "$(get-date) Info: Process Complete" -LogType "Info"

$objDetailLogFile.close()