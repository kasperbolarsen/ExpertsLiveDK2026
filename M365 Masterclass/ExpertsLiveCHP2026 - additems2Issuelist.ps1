 # this script will add items to the Issue%20tracker list in a site
 $siteUrl = "https://m365thinking.sharepoint.com/sites/ExpertsLive2026Test1"
$siteUrl = "https://m365thinking.sharepoint.com/sites/ExpertsLive2026Test2"

 $PnPClientId = "5cb5d6b0-29ff-4959-905d-ffc1c8208efc"



$conn = Connect-PnPOnline -Url $siteUrl -Interactive -ClientId $PnPClientId -ErrorAction Stop -ReturnConnection
$SharePointAdminUrl = "https://m365thinking-admin.sharepoint.com/"

$adminConn = Connect-PnPOnline -Url $SharePointAdminUrl -Interactive -ClientId $PnPClientId -ErrorAction Stop -ReturnConnection

function AddItemsToTheListColumnIssueList 
{
    

$listName = "Issue tracker"

$fields = Get-PnPField -List $listName -Connection $conn

for($i=1; $i -le 10; $i++)
{
    $title = "Issue $i"
    $description = "Description for issue $i"
    #shuffle the priority
    $priorities = @("High", "Normal", "Low")
    $priority = Get-Random -InputObject $priorities

    #shuffle the status
    $statuses = @("Open", "In Progress", "Closed")
    $status = Get-Random -InputObject $statuses

    #shuffle the assigned to
    # $assignedTos = @("KasperLarsen@M365Thinking.onmicrosoft.com", "Adele@M365Thinking.onmicrosoft.com", "liam.ogrady@M365Thinking.onmicrosoft.com")
    $AssignedTo = "KasperLarsen@M365Thinking.onmicrosoft.com"
    #$AssignedTo = Get-Random -InputObject $assignedTos
    Add-PnPListItem -List $listName -Values @{"Title"=$title; "Description"=$description; "Priority"=$priority; "Status"=$status; "Assignedto0"=$AssignedTo} -Connection $conn
    Write-Host "Added item: $title"
}
param (
        OptionalParameters
    )
    
}
function AddItemsToTheContentTypeIssueList 
{   
    $listName = "Issue tracker based on Content Type"
    $fields = Get-PnPField -List $listName -Connection $conn

    for($i=1; $i -le 10; $i++)
    {
        $title = "Issue $i"
        $description = "Description for issue $i"
        #shuffle the priority
        $priorities = @("High", "Normal", "Low")
        $priority = Get-Random -InputObject $priorities

        #shuffle the status
        $statuses = @("Open", "In Progress", "Closed")
        $status = Get-Random -InputObject $statuses

        #shuffle the assigned to
        # $assignedTos = @("KasperLarsen@M365Thinking.onmicrosoft.com", "Adele@M365Thinking.onmicrosoft.com", "liam.ogrady@M365Thinking.onmicrosoft.com")
        $AssignedTo = "KasperLarsen@M365Thinking.onmicrosoft.com"
        #$AssignedTo = Get-Random -InputObject $assignedTos
        Add-PnPListItem -List $listName -Values @{"Title"=$title; "Comment"=$description; "Priority"=$priority; "IssueStatus"=$status; "AssignedTo"=$AssignedTo} -Connection $conn
        Write-Host "Added item: $title"
    }
}

function AddItemsToTheContentTypeIssueList2 
{   
    $listName = "Issues List using Content type"
    $fields = Get-PnPField -List $listName -Connection $conn

    for($i=1; $i -le 10; $i++)
    {
        $title = "Issue $($i+100)"
        $description = "Description for $title"
        #shuffle the priority
        $priorities = @("High", "Normal", "Low")
        $priority = Get-Random -InputObject $priorities

        #shuffle the status
        $statuses = @("Open", "In Progress", "Closed")
        $status = Get-Random -InputObject $statuses

        #shuffle the assigned to
        # $assignedTos = @("KasperLarsen@M365Thinking.onmicrosoft.com", "Adele@M365Thinking.onmicrosoft.com", "liam.ogrady@M365Thinking.onmicrosoft.com")
        $AssignedTo = "KasperLarsen@M365Thinking.onmicrosoft.com"
        #$AssignedTo = Get-Random -InputObject $assignedTos
        Add-PnPListItem -List $listName -Values @{"Title"=$title; "Comment"=$description; "Priority"=$priority; "IssueStatus"=$status; "AssignedTo"=$AssignedTo} -Connection $conn
        Write-Host "Added item: $title"
    }
}
function RecursiveGetTerms($term)
{
    if($term.Terms.Count -gt 0)
    {
        foreach($childTerm in $term.Terms)
        {
            $allTerms.Add($childTerm) | Out-Null
            RecursiveGetTerms $childTerm 
        }
    }
    
}
function GetRandomTermIdFromGeoRelationTermSet 
{
    param (
        [Parameter(Mandatory=$true)]
        $Connection
    )
    
    try {

        
        # Get the GeoRelation term set
        $termSet = Get-PnPTermSet -Identity "Geo Relations" -TermGroup "Geo Relations" -Connection $adminConn
        
        $allterms = New-Object System.Collections.Generic.List[Microsoft.SharePoint.Client.Taxonomy.Term]
        # Get all terms in the term set
        $terms = Get-PnPTerm -TermSet $termSet -TermGroup "Geo Relations" -Connection $adminConn -Recursive -IncludeChildTerms
        
        foreach ($term in $terms) 
        {

            $allterms.Add($term)
            RecursiveGetTerms $term
            Write-Host "Term: $($term.Name), ID: $($term.Id)"
        }
        # Select a random term and return its ID
        $randomTerm = Get-Random -InputObject $allterms
        
        Write-Host "Selected random term: $($randomTerm.Name) with ID: $($randomTerm.Id)"
        return $randomTerm.Id.Guid
    }
    catch {
        Write-Error "Error retrieving term from GeoRelation term set: $_"
        return $null
    }
}

function AddItemsToTheELDKIssueList 
{   
    $listName = "ELDK26Issues"
    $fields = Get-PnPField -List $listName -Connection $conn

    for($i=1; $i -le 10; $i++)
    {
        $title = "Issue $($i+100)"
        $description = "Description for $title"
        #shuffle the priority
        $priorities = @("High", "Normal", "Low")
        $priority = Get-Random -InputObject $priorities

        #shuffle the status
        $statuses = @("Open", "In Progress", "Closed")
        $status = Get-Random -InputObject $statuses

        #shuffle the assigned to
        # $assignedTos = @("KasperLarsen@M365Thinking.onmicrosoft.com", "Adele@M365Thinking.onmicrosoft.com", "liam.ogrady@M365Thinking.onmicrosoft.com")
        $AssignedTo = "KasperLarsen@M365Thinking.onmicrosoft.com"
        #$AssignedTo = Get-Random -InputObject $assignedTos
        $geoRelation = GetRandomTermIdFromGeoRelationTermSet -Connection $conn
        Add-PnPListItem -List $listName -Values @{"Title"=$title; "Comment"=$description; "Priority"=$priority; "IssueStatus"=$status; "AssignedTo"=$AssignedTo; "GeoRelations"= $geoRelation} -Connection $conn
        Write-Host "Added item: $title"
    }
}

#AddItemsToTheContentTypeIssueList
#AddItemsToTheContentTypeIssueList2
# AddItemsToTheListColumnIssueList
AddItemsToTheELDKIssueList