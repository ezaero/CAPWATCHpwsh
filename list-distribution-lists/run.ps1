# Input bindings are passed in via param block.
param($Timer)

Connect-ExchangeOnline -ManagedIdentity -Organization COCivilAirPatrol.onmicrosoft.com

$allDistLists = Get-DistributionGroup -ResultSize Unlimited
Write-Host "Total DLs" $allDistLists.count

$allDistLists | ForEach-Object {
    Write-Progress -activity "Processing $_.DisplayName" -status "$i out of $totalmbx completed"
}