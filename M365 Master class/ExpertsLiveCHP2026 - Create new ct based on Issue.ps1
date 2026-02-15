#ExpertsLiveCHP2026 - Create new ct based on the Issue ct
# This script creates a new content type "ELDK26Issues" based on the Issue content type
# and adds a taxonomy site column "Geo Relations" to it

# Variables
$contentTypeHubUrl = "https://m365thinking.sharepoint.com/sites/Contenttypehub"
$baseContentTypeName = "Issue"
$newContentTypeName = "ELDK26Issues"
$siteColumnName = "Geo Relations"
$termSetName = "Geo Relations"
$termGroupName = "Geo Relations"

try 
{
    # Connect to Content Type Hub
    Write-Host "Connecting to Content Type Hub: $contentTypeHubUrl" -ForegroundColor Cyan
    Connect-PnPOnline -Url $contentTypeHubUrl -Interactive
    
    # Step 1: Find the Issue content type
    Write-Host "`nSearching for '$baseContentTypeName' content type..." -ForegroundColor Cyan
    $issueContentType = Get-PnPContentType | Where-Object { $_.Name -eq $baseContentTypeName }
    
    if ($null -eq $issueContentType) 
    {
        Write-Host "ERROR: '$baseContentTypeName' content type not found!" -ForegroundColor Red
        exit
    }
    
    Write-Host "Found '$baseContentTypeName' content type (ID: $($issueContentType.Id))" -ForegroundColor Green
    
    # Step 2: Check if the new content type already exists
    Write-Host "`nChecking if '$newContentTypeName' content type already exists..." -ForegroundColor Cyan
    $existingContentType = Get-PnPContentType | Where-Object { $_.Name -eq $newContentTypeName }
    
    if ($null -ne $existingContentType) 
    {
        Write-Host "WARNING: '$newContentTypeName' content type already exists (ID: $($existingContentType.Id))" -ForegroundColor Yellow
        Write-Host "Using existing content type..." -ForegroundColor Yellow
        $newContentType = $existingContentType
    }
    else 
    {
        # Step 3: Create new content type based on Issue
        Write-Host "Creating new content type '$newContentTypeName' based on '$baseContentTypeName'..." -ForegroundColor Cyan
        $newContentType = Add-PnPContentType -Name $newContentTypeName -ParentContentType $issueContentType
        Write-Host "Content type '$newContentTypeName' created successfully (ID: $($newContentType.Id))" -ForegroundColor Green
    }
    
    # Step 4: Get the term set for the taxonomy field
    Write-Host "`nRetrieving term set '$termSetName' from term group '$termGroupName'..." -ForegroundColor Cyan
    $termSet = Get-PnPTermSet -TermGroup $termGroupName -TermSet $termSetName -ErrorAction Stop
    
    if ($null -eq $termSet) 
    {
        Write-Host "ERROR: Term set '$termSetName' not found in term group '$termGroupName'!" -ForegroundColor Red
        exit
    }
    
    Write-Host "Found term set '$termSetName' (ID: $($termSet.Id))" -ForegroundColor Green
    
    # Step 5: Check if site column already exists
    Write-Host "`nChecking if site column '$siteColumnName' exists..." -ForegroundColor Cyan
    $existingField = Get-PnPField -Identity $siteColumnName -ErrorAction SilentlyContinue
    
    if ($null -eq $existingField) 
    {
        # Create the taxonomy site column
        Write-Host "Creating taxonomy site column '$siteColumnName'..." -ForegroundColor Cyan
        
        $fieldXml = @"
<Field Type="TaxonomyFieldType" 
       DisplayName="$siteColumnName" 
       Name="$($siteColumnName.Replace(' ',''))" 
       Required="FALSE" 
       Group="Custom Columns">
</Field>
"@
        
        $field = Add-PnPFieldFromXml -FieldXml $fieldXml
        
        # Connect the field to the term set
        Set-PnPField -Identity $field.Id -Values @
        {
            TermSetId = $termSet.Id
            AnchorId = "00000000-0000-0000-0000-000000000000"
        }
        
        Write-Host "Site column '$siteColumnName' created and connected to term set" -ForegroundColor Green
        $fieldToAdd = $field
    }
    else 
    {
        Write-Host "Site column '$siteColumnName' already exists" -ForegroundColor Yellow
        $fieldToAdd = $existingField
    }
    
    # Step 6: Add the site column to the content type
    Write-Host "`nAdding site column '$siteColumnName' to content type '$newContentTypeName'..." -ForegroundColor Cyan
    
    # Check if field is already in the content type
    $ctFields = Get-PnPProperty -ClientObject $newContentType -Property Fields
    $fieldInCT = $ctFields | Where-Object { $_.InternalName -eq $fieldToAdd.InternalName }
    
    if ($null -eq $fieldInCT) 
    {
        Add-PnPFieldToContentType -Field $fieldToAdd -ContentType $newContentType
        Write-Host "Site column '$siteColumnName' added to content type '$newContentTypeName'" -ForegroundColor Green
    }
    else 
    {
        Write-Host "Site column '$siteColumnName' is already in content type '$newContentTypeName'" -ForegroundColor Yellow
    }
    
    # Summary
    Write-Host "`n========================================" -ForegroundColor Green
    Write-Host "SCRIPT COMPLETED SUCCESSFULLY!" -ForegroundColor Green
    Write-Host "========================================" -ForegroundColor Green
    Write-Host "Content Type: $newContentTypeName" -ForegroundColor White
    Write-Host "Content Type ID: $($newContentType.Id)" -ForegroundColor White
    Write-Host "Site Column: $siteColumnName (Optional, Single Selection)" -ForegroundColor White
    Write-Host "Term Set: $termSetName" -ForegroundColor White
    Write-Host "========================================" -ForegroundColor Green
    
}
catch 
{
    Write-Host "`nERROR: $($_.Exception.Message)" -ForegroundColor Red
    Write-Host $_.Exception.StackTrace -ForegroundColor Red
}
finally 
{
    # Disconnect
    Write-Host "`nDisconnecting from SharePoint..." -ForegroundColor Cyan
    Disconnect-PnPOnline
}
