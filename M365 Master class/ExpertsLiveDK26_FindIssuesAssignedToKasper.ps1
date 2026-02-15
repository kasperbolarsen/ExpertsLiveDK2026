# find any issue assigned Kasper in the site
$siteUrl = "https://m365thinking.sharepoint.com/sites/ExpertsLive2026Test1"
$PnPClientId = "5cb5d6b0-29ff-4959-905d-ffc1c8208efc"
$FieldName = "Assignedto0"
$user = "KasperLarsen@M365Thinking.onmicrosoft.com"

if(-not $conn)
{
    $conn = Connect-PnPOnline -Url $siteUrl -Interactive -ClientId $PnPClientId -ErrorAction Stop -ReturnConnection
}


$hits = @()
$allLists = Get-PnPList -Connection $conn
foreach($list in $allLists)
{
    # $list
    $listItems = Get-PnPListItem -List $list -Connection $conn -ErrorAction Stop
    foreach($item in $listItems)
    {
        # $item.FieldValues
        if($item[$FieldName] -and $item[$FieldName].Email -eq $user)
        {
            #Write-Host "Found issue assigned to $user $($item['Title']) in list $($list.Title)"
            $hits += [PSCustomObject]@{
                Title = $item['Title']
                List = $list.Title
                Url = "$($conn.Url)/Lists/$($list.Title)/DispForm.aspx?ID=$($item.Id)"
            }
        }
    }
}   
if($hits.Count -eq 0)
{
    Write-Host "No issues found assigned to $user"
}
else
{
    Write-Host "Total issues found assigned to $($user) $($hits.Count)"
    $hits | Format-Table -AutoSize
}