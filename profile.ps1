# add ps1 with custom functions
$global:CUSTOMFUNCTIONS="C:\PATH\TO\functions.ps1"

# define path to txt-file to record all powershell history
$global:POSHHISTORYPATH="C:\Users\$($env:USERNAME)\Documents\WindowsPowerShell\posh_history_$($ENV:ComputerName).txt"

# set to false if you do not want the CUSTOMFUNCTIONS loaded nor be connected to RemoteExchange
$loadFunctions=$true

Clear-Host

##########################################
####        Connect to Exchange       ####
##########################################
function Test-Command ($Command) {
	#source: https://blogs.technet.microsoft.com/samdrey/2017/12/17/how-to-load-exchange-management-shell-into-powershell-ise-2/
	try {
		Get-Command $command -ErrorAction Stop
		return $True
	} Catch [System.SystemException] {
		return $False
	}
}

# connecting to Exchange Server - IF $env:ExchangeInstallPath is found
if ($loadFunctions -and $env:ExchangeInstallPath -and !(Test-Command "Get-Mailbox")) {
	$CallEMS = ". '$($env:ExchangeInstallPath)\bin\RemoteExchange.ps1'; Connect-ExchangeServer -auto -ClientApplication:ManagementShell "
	Invoke-Expression $CallEMS
	$exStatusMessageForLoadingInfos="Connected and commands have been loaded."
} else {
	$exStatusMessageForLoadingInfos="Not connected."
	Write-Host "Not connecting to Exchange, ExchangeInstallPath not present."
}

# load custom functions
if ((Test-Path $CUSTOMFUNCTIONS) -and $loadFunctions) {
	Write-Host "Loading $CUSTOMFUNCTIONS"
	. $CUSTOMFUNCTIONS
}

function Get-Time {
	return $(get-date | foreach { $_.ToLongTimeString() } )
}

##########################################
########        Custom Prompt      #######
##########################################
function prompt
{
	# get duration of last command in ms
	$lastCommand = (Get-History)[-1]
    $duration = ($lastCommand.EndExecutionTime - $lastCommand.StartExecutionTime).TotalMilliseconds.toString().split(",")[0]
    $durationPrompt = $duration.PadLeft(6," ")

    # build custom prompt including the current time
	$nextId = ((get-history -count 1).Id + 1).ToString().PadLeft(4, "0") ;
	$workDir=(Get-Location).path
	Write-Host "[" -NoNewLine
	Write-Host $(Get-Time) -foreground Green -NoNewLine
	Write-Host "]" -NoNewLine
	Write-Host " " -NoNewLine
	Write-Host "[$($durationPrompt)ms]" -NoNewLine
	Write-Host " " -NoNewLine
	Write-Host ("[" + $nextId + "] ") -ForegroundColor Yellow -NoNewLine
	Write-Host $env:computername -ForegroundColor White -NoNewLine
	Write-Host ' >' -ForegroundColor White -NoNewLine
	" "
	
	# add last completed command to POSHHISTORYPATH - if it doesn't already exist
	$historyitem=Get-History -count 1
	if ($historyitem.ExecutionStatus -eq "Completed") {
		$dateStart=(Get-Date $($historyitem.StartExecutionTime) -Format "dd.MM.yyyy HH:mm:ss")
		$dateEnd=(Get-Date $($historyitem.EndExecutionTime) -Format "dd.MM.yyyy HH:mm:ss")
		$lineToAdd="$($historyitem.Id)`t$($historyitem.ExecutionStatus)`t$dateStart`t$dateEnd`t$($historyitem.CommandLine)"
		$lastLineAdded=Get-Content $poshHistoryPath -Last 1
		if ($lineToAdd -ne $lastLineAdded) {
			#only add last command if it doesn't equal the last item
			Add-Content -path $poshHistoryPath -Value $lineToAdd
		}
	}
}

##########################################
######   Return Infos after Loading  #####
##########################################
$languageMode=$ExecutionContext.SessionState.LanguageMode
Write-Host "#################################################" -ForeGroundColor green
Write-Host $status -foregroundcolor Green
Write-Host "Language Mode is set to: " -NoNewLine
if ($languageMode -ne "FullLanguage") {
	Write-Host $languageMode -foregroundcolor Yellow
} else {
	Write-Host $languageMode -foregroundcolor Green
}
$executionPolicy=Get-ExecutionPolicy
$user=$ENV:USERNAME
$machine=$ENV:COMPUTERNAME
$machineIP=(Test-Connection $machine -Count 1).IPV4Address.IPAddressToString
$logonServer=$ENV:LOGONSERVER.replace("\","")
$logonServerIP=(Test-Connection $logonServer -Count 1).IPV4Address.IPAddressToString

Write-Host "       Execution Policy: $executionPolicy"
Write-Host "           Current User: $user"
Write-Host "        Running Machine: $machine - $machineIP"
Write-Host "           Logon Server: $logonServer - $logonServerIP"
if (Test-Path $CUSTOMFUNCTIONS) {
	Write-Host "run to reload functions: . `$CUSTOMFUNCTIONS"
}
Write-Host "      VerbosePreference: $VerbosePreference"
Write-Host "        DebugPreference: $DebugPreference"
Write-Host "      WarningPreference: $WarningPreference"
Write-Host "  ErrorActionPreference: $ErrorActionPreference"
Write-Host "    Exchange Connection: $exStatusMessageForLoadingInfos"
Write-Host "#################################################" -ForeGroundColor green