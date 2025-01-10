# Function to extract Team ID from the Team link
function Get-TeamIdFromLink {
    param (
        [string]$TeamLink
    )
    # Extract the Team ID from the URL
    if ($TeamLink -match "groupId=([a-f0-9\-]+)") {
        return $matches[1]
    } else {
        throw "Invalid Microsoft Teams link. Please ensure it includes the 'groupId' parameter."
    }
}

# Prompt the user for the Microsoft Teams link
$TeamLink = Read-Host "Please enter the Microsoft Teams link"

# Extract the Team ID
try {
    $TeamId = Get-TeamIdFromLink -TeamLink $TeamLink
    Write-Output "Extracted Team ID: $TeamId"
} catch {
    Write-Error $_.Exception.Message
    exit 1
}

# Prompt the user for the path to the CSV file
$csvPath = Read-Host "Please enter the full path to your CSV file"

# Check if the CSV file exists
if (-not (Test-Path -Path $csvPath)) {
    Write-Error "The CSV file does not exist at the specified path."
    exit 1
}

# Install the MicrosoftTeams module if not already installed
if (-not (Get-Module -ListAvailable -Name MicrosoftTeams)) {
    Install-Module -Name MicrosoftTeams -Force -AllowClobber
}

# Import the MicrosoftTeams module
Import-Module MicrosoftTeams

# Connect to Microsoft Teams
Connect-MicrosoftTeams

# Import members from the CSV file
$Members = Import-Csv -Path $csvPath

# Initialize an array to store successfully added members
$SuccessList = @()

# Add each member to the Microsoft Team
foreach ($member in $Members) {
    Write-Output "Attempting to add $($member.Email) to the team..."
    try {
        Add-TeamUser -GroupId $TeamId -User $member.Email -Role Member -Verbose
        Write-Output "Successfully added $($member.Email) to the team."
        $SuccessList += $member.Email
    } catch {
        Write-Error "Failed to add $($member.Email) to the team. Error: $_"
    }
}

# Disconnect from Microsoft Teams
Disconnect-MicrosoftTeams

# Report the successfully added members
if ($SuccessList.Count -gt 0) {
    Write-Output "The following members were successfully added to the team:"
    $SuccessList | ForEach-Object { Write-Output $_ }
} else {
    Write-Output "No members were successfully added."
}