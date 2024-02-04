function New-MonthlyADOUserStories {
    [CmdletBinding()]
    [Alias('nmus')]
    param (
        # Provide the work item type
        [Parameter(Mandatory = $false, 
            ValueFromPipeline = $true,
            ValueFromPipelineByPropertyName = $true, 
            ValueFromRemainingArguments = $false, 
            ParameterSetName = 'ADO Work Item Creation')]
        [ValidateNotNull()]
        [ValidateNotNullOrEmpty()]
        [Alias('type', 'workitem', 'task', 'bug')] 
        $WorkItemType = 'User Story',

        [Parameter(Mandatory = $false,
            ParameterSetName = 'ADO Work Item Creation')] 
        $organization = "https://dev.azure.com/datosmic/",

        [Parameter(Mandatory = $false,
            ParameterSetName = 'ADO Work Item Creation')] 
        $Project = "datosmic-project1",

        [Parameter(Mandatory = $false,
            ParameterSetName = 'ADO Work Item Creation')] 
        $ParentId = "1"
    )
 
    begin {

        $outputLogFolderPath = "$env:OneDrive\Automation"

        if ( (Test-Path -PathType container $outputLogFolderPath) -eq $false ) {
            Write-Host "$(Get-Date): $($outputLogFolderPath) doesn't exists so continuing with creation of folder for log file"
            New-Item -ItemType Directory -Path $outputLogFolderPath
        }
        else {
            Write-Host "$(Get-Date): $($outputLogFolderPath) exists so continuing with creation of log file"
        }

        # Output log file for generating output into file
        $OutputLogFilePath = "$outputLogFolderPath\WorkItemCreationLogs_$( (Get-Date).ToString('MMMMdd') ).txt"

        Write-Output "`n ================================== Output Logs of $(Get-Date) Iteration ================================== `n" | Out-File -FilePath $OutputLogFilePath -Append

        Write-Output "Output Log Path: $($OutputLogFilePath) `n" | Out-File -FilePath $OutputLogFilePath -Append
        
        Write-Output "$(Get-Date): Determining Month Name `n" | Out-File -FilePath $OutputLogFilePath -Append
        
        $FullMonthName = (Get-UICulture).DateTimeFormat.GetMonthName((Get-Date).Month) + ' ' + ((Get-Date).Year)

        Write-Output "$(Get-Date): Determined Month name: $FullMonthName `n" | Out-File -FilePath $OutputLogFilePath -Append

        $TeamMapping = @{
            Sales         = @('Sales & Marketing', 'MiriamG')
            Operations    = @('Operations', 'NestorW')
            Manufacturing = @('Manufacturing', 'LeeG')
        }

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
                $Tags,

                [Parameter(Mandatory = $true,
                    ParameterSetName = 'ADO Work Item Creation')] 
                $organization,
    
                [Parameter(Mandatory = $true,
                    ParameterSetName = 'ADO Work Item Creation')] 
                $Project
            )
                
            Write-Output "$(Get-Date): *** Creating Work Item *** `n" | Out-File -FilePath $OutputLogFilePath -Append
                   
            Write-Output "$(Get-Date): [Work Item Creation] - Passed values in New-ADOWorkItem function: `nOrganization: $($organization) `nProject: $($Project) `nTitle: $($Title) `nWorkItemType: $($WorkItemType) `nAssignedTo: $($AssignedTo) `nParentID: $ParentID `nTags: $($Tags) `n" | Out-File -FilePath $OutputLogFilePath -Append

            $WorkItem = az boards work-item create `
                --org $organization `
                --project $project `
                --title $Title `
                --type $WorkItemType `
                --assigned-to $AssignedTo 
        
            Write-Output "$(Get-Date): Created Work Item: `n$WorkItem `n" | Out-File -FilePath $OutputLogFilePath -Append
        
            # Adding Relation to the Work Item
            if ($? -eq $true) {
            
                $WorkItemID = ($WorkItem | ConvertFrom-Json).Id
        
                Write-Output "$(Get-Date): [Adding Work Item Relation] - Passed values in New-ADOWorkItem function: `nWorkItemID: $($WorkItemID) `nTitle: $($Title) `nWorkItemType: $($WorkItemType) `nAssignedTo: $($AssignedTo) `nParentID: $ParentID `nTags: $($Tags) `n" | Out-File -FilePath $OutputLogFilePath -Append
            
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

    }
    
    process {
        
        $WorkItemDetails = @()
        foreach ($TeamName in $TeamMapping.Keys) {
          
            $OrgName = $TeamMapping.$TeamName[0]

            $OrgLeadAlias = $TeamMapping.$TeamName[1]

            $Title = "[$FullMonthName] $OrgName Efforts Tracker"
            
            Write-Output "$(Get-Date): Generated Values: `nFullMonthName: $FullMonthName `nOrgName:$($TeamMapping.$TeamName[0]) `nTitle: $Title `nOrgLeadAlias: $($TeamMapping.$TeamName[1]) `n" | Out-File -FilePath $OutputLogFilePath -Append
    
            $WorkItemDetails += New-ADOWorkItem -Title $Title -WorkItemType $WorkItemType -AssignedTo $OrgLeadAlias -ParentID $ParentId -Tags $OrgName -organization $organization -Project $Project
        }
    }
    
    end {
        return $WorkItemDetails
    }
}
