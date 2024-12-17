Import-Function Get-Sites

# Run logic for each Site
foreach($site in Get-Sites){
    # Get Media Virtual Folder Item
    $mediaFolderItem = Get-Item -Path "master:$($site.ItemPath)/Media" -ErrorAction SilentlyContinue
    if ($null -eq $mediaFolderItem) {
        Write-Output "No Media Virtual Folder found for site: $($site.Name)"
        return
    }

    Write-Output "Updating Media Virtual Folder Sort Order for $($site.Name)"
    Reset-ItemField -Item $mediaFolderItem -Name "__Sortorder" -IncludeStandardFields
}