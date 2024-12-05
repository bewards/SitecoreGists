<#
    .SYNOPSIS
    The Remove-ExcessItemVersions function is designed to manage and potentially remove excess versions of Sitecore items based on specified criteria.

    .PARAMETER contextItem
        The Sitecore item to start the version removal process.
    .PARAMETER versionLimit
        The maximum number of versions an item should have.
    .PARAMETER includeDescendants
        A boolean indicating whether to include descendant items in the version removal process.
    .PARAMETER includeChildren
        A boolean indicating whether to include child items in the version removal process.
    .PARAMETER reportOnlySkipDelete
        A boolean indicating whether to only report items without deleting versions.
    .PARAMETER noReport
        A boolean indicating whether to skip generating a report.
#>

function Remove-ExcessItemVersions {
    param(
        [Sitecore.Data.Items.Item] $contextItem,
        [int] $versionLimit,
        [bool] $includeDescendants,
        [bool] $includeChildren,
        [bool] $reportOnlySkipDelete,
        [bool] $noReport
    )

    Write-Log "Executing Version Removal ContextItem: $($contextItem.Name), versionLimit: $($versionLimit), includeDescendants: $($includeDescendants), reportOnly: $($reportOnlySkipDelete), noReport $($noReport)" 

    $items = @($item) | 
        Where-Object { $_.Versions.Count -gt $versionLimit } | 
        Initialize-Item |
        Sort-Object -Property @{Expression={$_.Versions.Count}; Descending=$true}
         
    if ($includeChildren) {
        $items = @($item) + @(($item.GetChildren())) | 
            Where-Object { $_.Versions.Count -gt $versionLimit } | 
            Initialize-Item |
            Sort-Object -Property @{Expression={$_.Versions.Count}; Descending=$true}
    }         

    if ($includeDescendants) {
        $items = @($item) + @(($item.Axes.GetDescendants())) | 
            Where-Object { $_.Versions.Count -gt $versionLimit } | 
            Initialize-Item |
            Sort-Object -Property @{Expression={$_.Versions.Count}; Descending=$true}
     }   
     
     $reportProps = @{
        Property = @(
            @{Name="ID"; Expression={$_.ID}},
            "DisplayName",
            @{Name="Versions"; Expression={$_.Versions.Count}},
            @{Name="Versions To Delete"; Expression={versionsToDelete($_)}},
            @{Name="Versions To Keep"; Expression={versionsToKeep($_)}},
            @{Name="Published Version"; Expression={publishedVersion($_)}},
            @{Name="Kept Version"; Expression={versionsRetainedCount($_)}},
            @{Name="Path"; Expression={$_.ItemPath}}
        )
        Title = "Items with more than $count versions"
        InfoTitle = "Sitecore recommendation: Limit the number of versions of any item to the fewest possible."
        InfoDescription = "The report shows all items that have more than <b>$count versions</b> and how many should be kept."
    }
       
    if ($noReport -eq 1) {
        foreach($item in $items) {
            versionsToDelete($item)
        }          
    }
    else {
        $items | Show-ListView @reportProps       
    }
}

$versionsToRemoveList = New-Object System.Collections.ArrayList
$versionsToKeepList = New-Object System.Collections.ArrayList

#Functions
function shouldRemoveVersion($itemVersion, $versionNumber, $limit) {
    #If the version is published we can skip the count and date checks as we want to keep it
    if(isPublished($itemVersion)) {return $false}
    #If "Keep Version" is checked we can skip
    if($item["Keep Version"]){return $false}
    # If number of version exceeds limit - Versions with 'Keep Version' Checkbox then 
    if($versionNumber -gt $limit){ return $true}

    return $false
}

function getVersionsForItem($item){
    $versions = @()
    [Sitecore.Data.Managers.ItemManager]::GetVersions($item) | ForEach-Object { $versions += Get-Item -Path $item.Paths.FullPath -Version $_.Number }
    $versionsDesc = $versions | Where-Object {!$_["Keep Version"]}| Sort-Object -Property Version 
    return $versionsDesc
}

function versionsToDelete($item){
    $itemVersions = getVersionsForItem($item)
    
    # Set limit of versions to take into account item versions where 'Keep Version' checkbox is checked
    $versionsRetainedCount = versionsRetainedCount($item)
    $limit = $versionLimit - $versionsRetainedCount

    $versionCount = 0
    $versionsToRemove = 0
    foreach($itemVersion in $itemVersions){
        $versionCount +=1
        if(shouldRemoveVersion -ItemVersion $itemVersion -VersionNumber $versionCount -Limit $limit){
            #add to list to download
            $versionDate =  [Sitecore.DateUtil]::IsoDateToDateTime($itemVersion["__Updated"])
            $versionsToRemoveList.Add([pscustomobject]@{ID="$($itemVersion.ID)";Name="$($itemVersion.Name)";Path="$($itemVersion.Paths.FullPath)";Version="$($itemVersion.Version)";LastUpdated="$($versionDate)"}) | out-null
            $versionsToRemove +=1
            
            #remove version?
            if($reportOnlySkipDelete -eq 0) {
                Write-Log "Removing Item Version: $($itemVersion.ID) - $($itemVersion.Name) - version: $($itemVersion.Version)..." | out-null
                $itemVersion| Remove-ItemVersion
            }
        }
    }
    return $versionsToRemove
}

function versionsToKeep($item){
    $versionsDesc = getVersionsForItem($item)

    # Set limit of versions to take into account item versions where 'Keep Version' checkbox is checked
    $versionsRetainedCount = versionsRetainedCount($item)
    $limit = $versionLimit + $versionsRetainedCount

    $versionCount = 0
    $versionsToRemove = 0
    foreach($version in $versionsDesc){
        $versionCount +=1
        $itemVersion = Get-Item -Path $item.Paths.FullPath -Version $version
        if(shouldRemoveVersion -ItemVersion $itemVersion -VersionNumber $versionCount -Limit $limit){
            $versionsToRemove +=1
        }
        else{
             #add to list to download
             $versionDate =  [Sitecore.DateUtil]::IsoDateToDateTime($itemVersion["__Updated"])
             $versionsToKeepList.Add([pscustomobject]@{ID="$($itemVersion.ID)";Name="$($itemVersion.Name)";Path="$($itemVersion.Paths.FullPath)";Version="$($itemVersion.Version)";LastUpdated="$($versionDate)"}) | out-null
        }
    }
    $versionsToKeep = $versionsDesc.Count - $versionsToRemove
    return $versionsToKeep
}

function versionsRetainedCount($item){
    $versions = @()
    [Sitecore.Data.Managers.ItemManager]::GetVersions($item) | ForEach-Object { $versions += Get-Item -Path $item.Paths.FullPath -Version $_.Number }
    $versionsDesc = $versions | Where-Object {$_["Keep Version"]} | Sort-Object -Property Version
    Write-Host "Version Kept Count: $($versionsDesc.Count)"
    return $versionsDesc.Count
}

function publishedVersion($item){
    $versionsDesc = getVersionsForItem($item)
    $versionsToRemove = 0
    $publishedVersion = ""
    foreach($version in $versionsDesc){
        $itemVersion = Get-Item -Path $item.Paths.FullPath -Version $version
        $published = isPublished($itemVersion)
        if($published){
            $publishedVersion = $version
            break
        }
    }
    return $publishedVersion
}

function isPublished($item) {
    $path = $item.ItemPath -replace ('/','\')
    $itemExists = Get-Item -Path "web:$($path)" -Version $item.Version
    
    if($itemExists) {
        return $true
    } else {
       return $false
    }
}