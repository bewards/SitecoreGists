<#
    .SYNOPSIS
        Update the piped hostname values for all Site Definitions in a Sitecore instance and publishes the changes.
    .DESCRIPTION
        Finds all SXA site's setting items under the /sitecore/content/Websites folder and updates the Hostname field based on the Database field value.
#>

# hostname contains all possible environment specific site hostnames for CM and CD
$fullHostnameValue = "cmstest.site.com|sctest.site.com|azure-test-cd-001.azurewebsites.net|azure-test-cm-001.azurewebsites.net"
$linkProviderCM = "siteLinkProvider"
$linkProviderCD = "siteCDLinkProvider"

$applyUpdates = $true

$tenant = Get-Item -Path "master:/sitecore/content/Websites"
$siteFolders = $tenant | Get-ChildItem | Where-Object {$_.TemplateName -eq "Site Folder"}
$allSites = $tenant | Get-ChildItem | Where-Object {$_.TemplateName -eq "Site"}
$siteFolders | ForEach-Object {
    $_ | Get-ChildItem | Where-Object {$_.TemplateName -eq "Site"} | ForEach-Object {
        $allSites += $_
    }
}

$allSiteDefinitions = @()
$allSites | ForEach-Object {
    $siteGroupingPath = "$($_.Paths.FullPath)/Settings/Site Grouping"
    Get-ChildItem -Path $siteGroupingPath | ForEach-Object {
        $allSiteDefinitions += $_
    }
}

$total = $allSiteDefinitions.Count
$index = 1
$pct = 0

$startTime = Get-Date
$allSiteDefinitions | ForEach-Object {
    $siteDef = $_
    $timeRemaining = "X"
    if ($index -gt 0) {
        $secs = ([TimeSpan] ((Get-Date) - $startTime)).TotalSeconds
        $timeRemaining = [int][Math]::Ceiling(($secs/$index) * ($total-($index-1)))
    }
    $pct = [int][Math]::Floor(($index-1)/$total * 100)
    Write-Progress -Activity "[$($index)/$($total)] Update $($siteDef.Name)" -Status "Time Remaining: $($timeRemaining) sec - $($pct)% Complete" -PercentComplete $pct;

    $updatedLinkProvider = $linkProviderCM
    if ($siteDef["Database"] -eq "web") {
        $updatedLinkProvider = $linkProviderCD
    }
    if ($applyUpdates -eq $true) {
        $siteDef.Editing.BeginEdit()
        $siteDef.HostName = $fullHostnameValue
        $siteDef.LinkProvider = $updatedLinkProvider
        $siteDef.Editing.EndEdit()
    } else {
        Write-Host "[TEST] Update Site Definition Values on $($siteDef.Paths.FullPath)"
    }
    $index++
}
if ($applyUpdates -eq $true) {
    # Publish all updated items
    $database = Get-Database -Name "master"
    $targets = @()
    foreach($publishingTarget in [Sitecore.Publishing.PublishManager]::GetPublishingTargets($database)) {
        $targets += Get-Database -Name $publishingTarget[[Sitecore.FieldIDs]::PublishingTargetDatabase]
    }
    $languages = [Sitecore.Data.Managers.LanguageManager]::GetLanguages($database)
    $allSiteDefinitions | Where-Object {$_["Database"] -eq "web"} | ForEach-Object {
        [Sitecore.Publishing.PublishManager]::PublishItem($_,$targets,$languages,$true,$true,$false)
    }
}

$totalTime = "{0:hh\:mm\:ss\.fff}" -f ([TimeSpan] ((Get-Date) - $startTime))
Write-Host "Completed in $($totalTime) time"

