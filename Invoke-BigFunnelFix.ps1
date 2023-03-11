<# This script is designed to get mailboxes on an Exchange Server with Markers of a Big Funnel Search issue
You can find out more here https://www.reddit.com/r/exchangeserver/comments/wpfq14/big_funnel_indexing_workaround_on_exchange_2019/ #> 

param (
	[switch]$verbose,
	[switch]$TriggerAssistant,
	[switch]$WriteCSV
)

Add-PSSnapin Microsoft.Exchange.Management.PowerShell.SnapIn; 

$AllExchServers = Get-ExchangeServer
$SysName = $env:computername

if(-not ($AllExchServers.Name -contains $sysname)){
	Write-Host "Not running on a local Exchange server!!!"
	Write-host "This Script MUST be run on a local server. It can not be run remotely because of MS issues with proxy commands"
	Write-host "At the time of implementation, the command to initiate the run doesn't support remote servers"
	exit
	exit
}else{

	$CurTime = Get-Date -Format "yyyy-MM-dd HH.mm.ss"
	$OutputFile = "D:\BigFunnel\$SysName-BigFunnelReport-$CurTime.csv"
	$FolderName = "D:\BigFunnel\"

	Write-Host -foregroundcolor Blue "Getting mailbox info. Please hold."
	$AllMbxOnServer = Get-mailbox -Server $SysName -ResultSize unlimited

	[System.Collections.ArrayList]$CombinedInfo = @()

	Foreach ($Mbx in $AllMbxOnServer){
		$MbxStats = Get-MailboxStatistics -Identity $Mbx.Alias
		$MbxNIC = $MbxStats.BigfunnelNotIndexedCount
		
		If ($MbxNIC -ge "1"){
			
			Write-Host -foregroundcolor red "$Mbx.Alias has markers of BigFunnelIssues ($MbxNIC items are unindexed )"
			
			if($triggerassistant){
				Write-Host -foregroundcolor blue "Triggering Mailbox Assistant"
				Start-MailboxAssistant -Identity $Mbx.Alias -AssistantName BigFunnelRetryFeederTimeBasedAssistant
			}
			
			$ThisCombinedInfo = [PsCustomObject]@{
					MbxAlias							= $mbx.Alias
					DisplayName							= $MbxStats.DisplayName
					Mailboxname 						= $MBX.Database.name
					MailboxActiveServer 				= $MBX.ServerName
					BigFunnelMessageCount				= $MbxStats.BigFunnelMessageCount
					BigFunnelIndexedCount				= $MbxStats.BigFunnelIndexedCount
					BigFunnelPartiallyIndexedCount	 	= $MbxStats.BigFunnelPartiallyIndexedCount
					BigFunnelNotIndexedCount			= $MbxStats.BigFunnelNotIndexedCount
					BigFunnelCorruptedCount 			= $MbxStats.BigFunnelCorruptedCount
					BigFunnelStaleCount 				= $MbxStats.BigFunnelStaleCount
					BigFunnelShouldNotBeIndexedCount 	= $MbxStats.BigFunnelShouldNotBeIndexedCount
				}
				$CombinedInfo.Add($ThisCombinedInfo) | out-null
		}else{
			if($verbose){
				Write-Host -foregroundcolor green "$Mbx.Alias has no markers of BigFunnelIssues ($MbxNIC items are unindexed )"
			}
		}
	}

	Write-host "The following were detected as having issues and triggered the fix"
	$CombinedInfo | Format-Table -autosize

	if($WriteCSV){
		if (!(Test-Path $FolderName)){
			Write-host "Creating directory $Foldername for logs"
			New-Item $FolderName -ItemType Directory
		}
		Write-Host "CSV Written to $OutputFile"
		$CombinedInfo | Export-csv $OutputFile -NoType
	}
}
