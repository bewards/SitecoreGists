<#
    .SYNOPSIS
        Updates the Default Page Design for all SXA sites with a new Header rendering.
    .DESCRIPTION
        Finds every Site's Default Page Design and updates the Header rendering field to a new value.
#>

function Update-PageDesign {
    param (
        $site
    )
    Write-Log "Updating Page Design for $($site.Name)"
    
    $pageDesign = Get-Item -Path "master:$($site.ItemPath)/Presentation/Page Designs/Default"

    [Sitecore.Text.ListString]$fieldIds = $pageDesign.Fields["0966B999-0D0E-4278-ACC9-9DA69D461FE6"].Value

    $newIds = @()
    $newIds += $oldHeader 
    foreach ($fieldId in $fieldIds)
    {
        if (($fieldId -ne $oldHeader) -and ($fieldId -ne $newHeader))
        {
            $newIds += $fieldId
        }
    }

    $ids = [System.String]::Join("|", $newIds)
    Write-Log "Modifying field to $($ids)"
    $pageDesign.Editing.BeginEdit()
    $pageDesign["0966B999-0D0E-4278-ACC9-9DA69D461FE6"] = $ids
    $pageDesign.Editing.EndEdit()

    Write-Log "Publishing item..."
    Publish-Item -Item $pageDesign -PublishMode Smart -Recurse -CompareRevisions -AsJob
}

$applyChanges = $true
$verbose = $false
if ($verbose -eq $true){ $VerbosePreference = "Continue"} else { $verbosePreference = "SilentlyContinue"}

Write-Log "$(Get-Date -Format o) - Start - Replace with Regular Header for $((Get-User -Current).Name)"

# Get all Sites
$sites = Get-ChildItem -Path "master:/sitecore/content/websites"

$oldHeader = "{6D618591-A9B0-45F6-B67A-A76B1AA5BDAA}"
$newHeader = "{C799A584-BBE6-4EC9-9FE3-DA201E14F38C}"
# Run logic for each Site
foreach($site in $sites){
    if($site.Id -eq "{964F7521-036E-4FA4-A690-21F69DE92B83}") # if site is search site, skip
    {
        continue
    }
    
    if ($site.TemplateName -eq "Site Folder")
    {
        $groupSites = Get-ChildItem -Item $site | Where-Object {$_.Id -ne $globalSite.Id}
        
        foreach($groupSite in $groupSites)
        {
            Update-PageDesign -Site $groupSite
        }
    }
    else
    {
        Update-PageDesign -Site $site
    }
}

Write-Log "$(Get-Date -Format o) - End - Replace with Regular Header for $((Get-User -Current).Name)"