[CmdletBinding(DefaultParameterSetName = 'MailboxId')]
param (
    [Parameter(Mandatory, ParameterSetName = 'MailboxId')]
    [string]
    $MailboxId,

    [Parameter(Mandatory, ParameterSetName = 'All')]
    [switch]
    $All
)

if ($PSVersionTable.PSEdition -eq 'Core') {
    $PSStyle.Progress.View = 'Classic'
}

$ProgressPreference = 'Continue'

if ($PSCmdlet.ParameterSetName -eq 'MailboxId') {
    "Getting mailbox ($($MailboxId))" | Out-Default
    $mailbox = Get-Mailbox -Identity $MailboxId
}
if ($PSCmdlet.ParameterSetName -eq 'All') {
    "Getting all mailbox..." | Out-Default
    $mailbox = Get-Mailbox -ResultSize Unlimited
}

if (!$mailbox) {
    # If no mailbox was retrieve for whatever reason.. terminate the script.
    return $null
}

$total = $mailbox.Count
$counter = 0
$pctComplete = 0

"Mailbox count: $($total)" | Out-Default

$result = @()
"Getting mailbox statistics..." | Out-Default
$mailbox | ForEach-Object {
    $currentMailbox = $_
    $counter++
    $pctComplete = ($counter / $total) * 100
    Write-Progress -Activity "Getting mailbox statistics... ($([math]::Round($pctComplete,2))%)" -Status "[$($counter) / $($total)] [$($currentMailbox.DisplayName)]"

    $stats = Get-MailboxStatistics -Identity $currentMailbox.ExchangeGuid
    $stats.TotalItemSize = (($stats.TotalItemSize.ToString() -split '\(')[1] -replace ' bytes\)', '' -replace ',', '') -as [Int64]
    $stats | Add-Member -MemberType NoteProperty -Name PrimarySmtpAddress -Value $currentMailbox.PrimarySmtpAddress


    $result += $stats
}

$result | Add-Member -MemberType ScriptProperty -Name TotalItemSizeMB -Value { [math]::round(($this.TotalItemSize / 1MB), 2) }
$result | Add-Member -MemberType ScriptProperty -Name TotalItemSizeGB -Value { [math]::round(($this.TotalItemSize / 1GB), 2) }
$result