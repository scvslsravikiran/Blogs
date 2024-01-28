function New-ADOWorkItem {
    param (
        [Parameter(Mandatory = $true,
            ParameterSetName = 'ADO Work Item Creation')] 
        $Title,
        
        [Parameter(Mandatory = $true,
            ParameterSetName = 'ADO Work Item Creation')] 
        $WorkItemType,
        
        [Parameter(Mandatory = $true,
            ParameterSetName = 'ADO Work Item Creation')] 
        $AssignedTo,
        
        [Parameter(Mandatory = $true,
            ParameterSetName = 'ADO Work Item Creation')] 
        $ParentID,

        [Parameter(Mandatory = $true,
            ParameterSetName = 'ADO Work Item Creation')] 
        $Tags
    )
    

    $organization = "https://dev.azure.com/datosmic/"
    $project = "datosmic-project1"
        
    $outputLogFolderPath = "$env:OneDrive\Automation"

    if ( (Test-Path -PathType container $outputLogFolderPath) -eq $false ) {
        Write-Host "$(Get-Date): $($outputLogFolderPath) doesn't exists so continuing with creation of folder for log file"
        New-Item -ItemType Directory -Path $outputLogFolderPath
    }
    else {
        Write-Host "$(Get-Date): $($outputLogFolderPath) exists so continuing with creation of log file"
    }

    # Output log file for generating output into file
    $OutputLogFilePath = "$outputLogFolderPath\TasksCreationLogs_$( (Get-Date).ToString('MMMMdd') ).txt"


    Write-Output "$(Get-Date): *** Creating Work Item ***`n" | Out-File -FilePath $OutputLogFilePath -Append
                
    $WorkItem = az boards work-item create `
        --org $organization `
        --project $project `
        --title $Title `
        --type $WorkItemType `
        --assigned-to $AssignedTo 

    Write-Verbose "$(Get-Date): Created Work Item: `n$WorkItem"

    # Adding Relation to the Work Item
    if ($? -eq $true) {
        
        $WorkItemID = ($WorkItem | ConvertFrom-Json).Id

        Write-Output "$(Get-Date): Passed values in New-ADOWorkItem function: `nWorkItemID: $($WorkItemID) `nTitle: $($Title) `nWorkItemType: $($WorkItemType) `nAssignedTo: $($AssignedTo) `nParentID: $ParentID `nTags: $($Tags) `n" | Out-File -FilePath $OutputLogFilePath -Append
        
        $WorkItemDetails = az boards work-item relation add `
            --org $organization `
            --id $WorkItemID  `
            --relation-type parent `
            --target-id $ParentID  

        if ($? -eq $true) {   
            Write-Output "$(Get-Date): Added $ParentID User Story as a Parent for $WorkItemID Task... `nWorkItemDetails: $WorkItemDetails `n " | Out-File -FilePath $OutputLogFilePath -Append
            
            $Tags = $Tags

            $Task = az boards work-item update `
                --org $organization `
                --fields System.Tags="$Tags" `
                --id $WorkItemID  

            if ($? -eq $true) { 
                Write-Output "$(Get-Date): Added Tags: $($Task.fields.'System.Tags') to the Task $($Task.Id)..`nTask Details after adding Tags: $Task `n" | Out-File -FilePath $OutputLogFilePath -Append

                $Task = $Task | ConvertFrom-Json

                Write-Host "$(Get-Date): Successfully created `
                Title: $($Task.fields.'System.Title') `
                WorkItem Id: $($Task.Id) `
                WorkItem URL: $($Task.url) `
                ParentID: $ParentID `
                Relation Type: $($Task.relations.attributes.name) `
                Related WorkItem URL: $($Task.relations.url) `
                WorkItem Type: $($Task.fields.'System.WorkItemType') `
                IterationPath:  $($Task.fields.'System.IterationPath') `
                State: $($Task.fields.'System.State') `
                Tags: $($Task.fields.'System.Tags') `
                AssignedTo: $($Task.fields.'System.AssignedTo'.displayName) `n`n" -ForegroundColor Cyan
            
                Write-Output "$(Get-Date): Successfully created `
                Title: $($Task.fields.'System.Title') `
                WorkItem Id: $($Task.Id) `
                WorkItem URL: $($Task.url) `
                ParentID: $ParentID `
                Relation Type: $($Task.relations.attributes.name) `
                Related WorkItem URL: $($Task.relations.url) `
                WorkItem Type: $($Task.fields.'System.WorkItemType') `
                IterationPath:  $($Task.fields.'System.IterationPath') `
                State: $($Task.fields.'System.State') `
                Tags: $($Task.fields.'System.Tags') `
                AssignedTo: $($Task.fields.'System.AssignedTo'.displayName) `n`n" | Out-File -FilePath $OutputLogFilePath -Append

                return $Task
            }
            else {
                Write-Error "$(Get-Date): Work Item creation failed with error at Task."
            }
        }
        else {
            Write-Error "$(Get-Date): Work Item creation failed with error at WorkItemDetails"
        }
    }
    else {
        Write-Error "$(Get-Date): Work Item creation failed with error at WorkItem"
    }
}

$Title = "Created Work Item with Automation"
$WorkItemType = "Task"
$AssignedTo = "AdeleV"
$ParentId = 3
$Tags = "Automation;Task"

$Task = New-ADOWorkItem -Title $Title -WorkItemType $WorkItemType -AssignedTo $AssignedTo -ParentID $ParentId -Tags $Tags
