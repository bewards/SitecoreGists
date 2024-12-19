<#
    .SYNOPSIS
        Finds latest version of items where the display name contains ASCII encoding and optionally (triple) decodes it.
    .DESCRIPTION
        Item display names can get double or triple encoded during a migration or import process. This script will find items with display names that contain ASCII encoding and optionally (triple) decode it back to the original value with ASCII codes (&#amp; etc.)
#>

# APPLY UPDATES - set to $true to update the display names with triple decoded values fix
$applyUpdates = $false

<#
.SYNOPSIS
The purpose of escaping the display name is for in the ListView, as a single ASCII (&#x27;) doesn't show in list view without escaping with &/. This function is only used for display purposes.

.PARAMETER string
The value.
#>
function Set-EscapeCharacters {
    Param(
        [parameter(Mandatory = $true, Position = 0)]
        [String]
        $string
    )
    $string = $string -replace '\&', '&/'
    $string
}

# PREDICATE - root
$root = Get-Item -Path "master:" -ID "{829352E7-5D13-4AF1-B877-1048D55D42B8}"   # root (Websites tenant)
$criteriaRoot = @{Filter = "DescendantOf"; Value = $root }
$predicateRoot = New-SearchPredicate -Criteria $criteriaRoot

# PREDICATE - latest version
$criteriaLatest =  @{Filter = "Equals"; Field = "_latestversion"; Value = "1"}
$predicateLatest = New-SearchPredicate -Criteria $criteriaLatest

# PREDICATE COMBINE - root AND latest
$predicateLatestRoot = New-SearchPredicate -First $predicateRoot -Second $predicateLatest -Operation And

# PREDICATE - display name
# ASCII CODES (minus the end semi-colon so search doesn't translate it)
$criteriaDisplayName = @(
    # DECIMAL CODES
    @{Filter = "Equals"; Field = "_displayname"; Value = "&#33"; },
    @{Filter = "Equals"; Field = "_displayname"; Value = "&#34"; },
    @{Filter = "Equals"; Field = "_displayname"; Value = "&#35"; },
    @{Filter = "Equals"; Field = "_displayname"; Value = "&#36"; },
    @{Filter = "Equals"; Field = "_displayname"; Value = "&#37"; },
    @{Filter = "Equals"; Field = "_displayname"; Value = "&#38"; },
    @{Filter = "Equals"; Field = "_displayname"; Value = "&#40"; },
    @{Filter = "Equals"; Field = "_displayname"; Value = "&#41"; },
    # HEX CODES
    @{Filter = "Equals"; Field = "_displayname"; Value = "&#x2c"; },
    @{Filter = "Equals"; Field = "_displayname"; Value = "&#x27"; },
    @{Filter = "Equals"; Field = "_displayname"; Value = "&#x28"; },
    @{Filter = "Equals"; Field = "_displayname"; Value = "&#x29"; },
    #CHAR CODES
    @{Filter = "Equals"; Field = "_displayname"; Value = "&amp"; },
    @{Filter = "Equals"; Field = "_displayname"; Value = "&lt"; },
    @{Filter = "Equals"; Field = "_displayname"; Value = "&gt"; },
    @{Filter = "Equals"; Field = "_displayname"; Value = "&quot"; },
    @{Filter = "Equals"; Field = "_displayname"; Value = "&semi"; }
)
$predicateDisplayName = New-SearchPredicate -Operation Or -Criteria $criteriaDisplayName

# PREDICATE COMBINE - AND
$predicate = New-SearchPredicate -First $predicateLatestRoot -Second $predicateDisplayName -Operation And

$props = @{
    Index = "sitecore_master_index"
    WherePredicate = $predicate
}

$searchItems = Find-Item @props
$filteredItems = @()

## FILTER afterwards - search has false positives
$searchItems | % {
    $item = $_
    $dName = $item.Fields["_displayname"] # item type is still SearchResultItem, so will need to use Fields collection (otherwise have to convert each one with Initialize-Item)
    
    if([string]::IsNullOrEmpty($dName)) {
        continue
    }
    
    if($dName -match '(&#33|&#34|&#35|&#36|&#37|&#38|&#40|&#41|&#x2c|&#x27|&#x28|&#x29|&amp|&lt|&gt|&quot|&semi)') {
        $filteredItems += $item
        # only enable this if you need full access to the item
        # if (($item | Initialize-Item).Versions.IsLatestVersion()) {
        #     $filteredItems += $item    
        # }

        # APPLY UPDATES - triple decode item display names
        if ($applyUpdates) {
            $item.Editing.BeginEdit()
            $item.Fields["_displayname"].Value = [System.Web.HttpUtility]::HtmlDecode([System.Web.HttpUtility]::HtmlDecode([System.Web.HttpUtility]::HtmlDecode($dName)))
            $item.Editing.EndEdit()
        }
    }
}

$reportProps = @{
    Title = "Item Display Names with ASCII Encoding"
    PageSize = 50
}
$filteredItems | Show-ListView @reportProps -Property `
    @{ Label = "Display Name"; Expression = { Set-EscapeCharacters $_.Fields["_displayname"] } },
    @{ Label = "Display Name Triple Decoded"; Expression = { [System.Web.HttpUtility]::HtmlDecode([System.Web.HttpUtility]::HtmlDecode([System.Web.HttpUtility]::HtmlDecode($_.Fields["_displayname"]))) } },
    @{ Label = "Name"; Expression = { $_.Name } },
    @{ Label = "Template"; Expression = { $_.TemplateName } },
    @{ Label = "Version"; Expression = { $_.Version } }
    
## DEBUG AREA - item 1 has ascii issues, item 2 only has a single ascci not converted (&#x27;), but it doesn't show in list view without escaping with &/

# $item1 = Get-Item -Path "master:" -ID "{0028CB5B-2D78-4B8A-AD71-CA89383A792E}"
# $d1 = $item1."__Display Name"
# $d1
# # [System.Web.HttpUtility]::HtmlDecode($d1)

# $item2 = Get-Item -Path "master:" -ID "{34BC6F10-EB7A-4858-BF0B-291D117AE63D}"
# $d2 = $item2."__Display Name"
# $d2
# # [System.Web.HttpUtility]::HtmlDecode($d2)
# Set-EscapeCharacters $d2