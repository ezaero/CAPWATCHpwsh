using namespace System.Net

# Input bindings are passed in via param block.
param($Request, $TriggerMetadata)

Connect-ExchangeOnline -ManagedIdentity -Organization COCivilAirPatrol.onmicrosoft.com

$allDistLists = Get-DistributionGroup -ResultSize Unlimited
Write-Host 'Total DLs' $allDistLists.count

$allDistLists | ForEach-Object {
    Write-Progress -Activity "Processing $_.DisplayName" -Status "$i out of $totalmbx completed"
}