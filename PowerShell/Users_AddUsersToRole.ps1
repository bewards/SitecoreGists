<#
    .SYNOPSIS
        Adds Sitecore users to a specified role using either a string of email users or a user input.
    .DESCRIPTION
        Finds every Site's Default Page Design and updates the Header rendering field to a new value.
#>

#-----------------Start - Modal Configuration ---------------#
$title = "Add Users to Role"
$description = "Add Users to Role by Email String Array or using the User Input"
$icon = "Office/32x32/users4.png"
$currentUser = (Get-User -Current).Name

Write-Log "$(Get-Date -Format o) - Start - Executing Script $($title) for $($currentUser)"
#-----------------End - Modal Configuration -----------------#

#-----------------Start - Action Tab Configuration ----------#
$actionTab = @(
    @{ Name = "userStringInput"; Value=""; Title="Email Users String Input"; lines=3; Tooltip="Enter string of email users"; Tab="Action"; Mandatory=$false; Placeholder="LastName, Michael <LastNameM@site.com>; LastName, Sam <LastNameS@site.com>;"; ParentGroupId=1; }, 
    @{ Name = "userInput"; Value=""; Title="Select Users"; Tooltip="Select Users to Add to Role"; Editor="user multiple"; Tab="Action"; }, 
    @{ Name = "roleInput"; Mandatory=$true; Title="Select Role"; Tooltip="Select Role to Add to Users"; Editor="role multiple"; Domain="sitecore"; Tab="Action"; }
)
#-----------------End - Action Tab Configuration -------------#

#-----------------Start - Options Tab Configuration ----------#
$optionsTab = @(
    @{ Name = "applyUpdates"; Value=$false; Title="Apply Updates"; Tooltip="Apply configured changes"; Editor="checkbox"; Tab="Options"; Columns = 4; GroupId=1; }
    @{ Name = "Description"; Title="Description"; Value=$description; editor="info"; Tab="Options"; ParentGroupId=1; }
)
#-----------------End - Options Tab Configuration ------------#

#-----------------Start - Apply Configuration ----------------#
$allTabs = $optionsTab + $actionTab

$dialogProps = @{
    Title = $title
    Description = $description
    Width = 650 
    Height = 700
    OkButtonName = "Continue"
    CancelButtonName = "Cancel"
    ShowHints = $true
    Icon = $icon
    Parameters = $allTabs
}

#-----------------End - Apply Configuration -------------------#

#-----------------Start - Run Modal(s) ------------------------#

$result = Read-Variable @dialogProps

if($result -ne "ok") {
    Write-Log "$(Get-Date -Format o) - Cancelled - Executing Script $($title) for $($currentUser)"
    Exit
}

if($applyUpdates -eq $true){
    # If user chose to apply updates, then ask to confirm choice
    $confirmation = Show-ModalDialog -Control "ConfirmChoice" -Parameters @{btn_0="Execute"; btn_1="Cancel";te="Please confirm choice to apply updates"} -Height 120 -Width 500

    if($confirmation -ne "btn_0"){
        Write-Log "$(Get-Date -Format o) - Cancelled - Executing Script $($title) for $($currentUser)"
        Exit
    }
}

#-----------------End - Run Modal(s) --------------------------#

#-----------------Start - Execute Action ----------------------#
$userData = @()
$missingUserData = @()

# Parse User String Input
if($userStringInput -ne "" -and $null -ne $userStringInput){
    $usernames = @()
    #split into an array:
    $userStringInput.split(';') | ForEach-Object {
        $emailStr = $_
        $l = $emailStr.IndexOf("<") + 1
        if ($l -eq -1) {
            Write-Host "no value for L setting to 0"
            $l = 0
        }
        $r = $emailStr.IndexOf("@")
        $u = $emailStr.Substring($l,$r-$l)
        $userNames += $u
    }
    #We have the usernames...
    $userNames | ForEach-Object {
        $userName = $_
        $user = Get-User -Filter "sitecore\$($userName)"
        if ($user -ne $null) {
            $userData += $user
        } else {
            $missingUserData += $userName
        }
    }
}

# Parse User Input
if($userInput.Count -gt 0 -and $null -ne $userInput -and $userInput -ne "" ){
    foreach ($user in $userInput) {
        $userData += Get-User $user
    }
}

if($applyUpdates -eq $true){
    if($userData.Count -eq 0){
        Write-Host "No Users Provided - See Missing Users output" -ForegroundColor Red
    }
    elseif($roleInput.Count -eq 0){
        Write-Host "No Role Provided - No Changes Applied" -ForegroundColor Red
    }
    else {
        foreach ($role in $roleInput) {
            Write-Host "Adding Role: '$($role)' to Following Users " -ForegroundColor Green
            $userData | Format-Table -AutoSize
            Write-Log "Adding Role: '$($role)' to Following Users $($userData) Request"
            try {
                Add-RoleMember -Identity $role -Members $userData
            }
            catch {
                Write-Host "Unable to Add Role: '$($role)' to Following Users " -ForegroundColor Green
                $userData | Format-Table -AutoSize
                Write-Log "$(Get-Date -Format o) - Error - Executing Script $($title) for $($currentUser)" -Log Error
                Write-Log "$(Get-Date -Format o) - Error - Adding Role: $($role) to Following Users"
                Write-Log "$(Get-Date -Format o) - Error - Users: $($userData -join ',')"
            }
            
        }
    }
}

if($applyUpdates -eq $false){
    Write-Host "Apply Updates checkbox was not clicked, see below for a list of changes that would be applied" -ForegroundColor Green
    Write-Host "Re-Run Script with 'Apply Updates' checkbox clicked to apply the below changes" -ForegroundColor Green
    foreach ($role in $roleInput) {
        Write-Host "Adding Role: '$($role)' to Following Users"
        $userData | Format-Table -AutoSize
    }
}

if($missingUserData.Count -gt 0){
    Write-Host "########Missing Users########" -ForegroundColor Red
    Write-Host "Unable to find following users" -ForegroundColor Red
    $missingUserData | Format-Table -AutoSize
}

Show-Result -Text

Write-Log "$(Get-Date -Format o) - Finished - Executing Script $($title) for $($currentUser)"
#-----------------End - Execute Action ----------------------#