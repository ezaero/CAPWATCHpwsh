function Get-RegionalDistributionGroups {
    return @(
        [PSCustomObject]@{
            Name = "Group 1 - Northern Colorado"
            Alias = "group1"
            EmailAddress = "group1@cowg.cap.gov"
            Units = @("072", "191", "099", "136", "147", "022", "068")
        },
        [PSCustomObject]@{
            Name = "Group 2 - Western Slope"
            Alias = "group2"
            EmailAddress = "group2@cowg.cap.gov"
            Units = @("053", "015", "189", "141", "181")
        },
        [PSCustomObject]@{
            Name = "Group 3 - Southern Colorado"
            Alias = "group3"
            EmailAddress = "group3@cowg.cap.gov"
            Units = @("805", "098", "030", "080", "159", "807")
        },
        [PSCustomObject]@{
            Name = "Group 4 - Central Colorado"
            Alias = "group4"
            EmailAddress = "group4@cowg.cap.gov"
            Units = @("143", "157", "148", "162", "183", "031", "163", "186", "173")
        }
    )
}

function Get-RegionalDistributionGroupMembers {
    param (
        [Parameter(Mandatory = $true)]
        [PSCustomObject]$Group,

        [Parameter(Mandatory = $true)]
        [array]$AllUsers
    )

    $companyNames = $Group.Units | ForEach-Object { "$($env:WING_DESIGNATOR)-$_" }
    return @(
        $AllUsers |
            Where-Object { $companyNames -contains $_.companyName } |
            Select-Object -ExpandProperty mail |
            Where-Object { -not [string]::IsNullOrWhiteSpace($_) } |
            Sort-Object -Unique
    )
}
