function New-WeeklyADOWorkItems {
    [CmdletBinding()]
    [Alias('nwado')]
    param (
        # Provide the work item type
        [Parameter(Mandatory = $false, 
            ValueFromPipeline = $true,
            ValueFromPipelineByPropertyName = $true, 
            ValueFromRemainingArguments = $false, 
            ParameterSetName = 'Azure Parameter Set')]
        [ValidateNotNull()]
        [ValidateNotNullOrEmpty()]
        [Alias('type', 'workitem', 'task', 'bug')] 
        $WorkItemType = 'Task',

        # Provide the organization
        [Parameter(Mandatory = $false,
            ParameterSetName = 'Azure Parameter Set')]
        $Organization = "https://dev.azure.com/datosmic/",

        # Provide the organization
        [Parameter(Mandatory = $false,
            ParameterSetName = 'Azure Parameter Set')] 
        $Project =  "datosmic-project1"
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

        Write-Host "$(Get-Date): Output Log Path: $($OutputLogFilePath) `n" 

        Write-Output "`n ================================== Output Logs of $(Get-Date) Iteration ================================== `n" | Out-File -FilePath $OutputLogFilePath -Append

        Write-Output "$(Get-Date): Output Log Path: $($OutputLogFilePath) `n" | Out-File -FilePath $OutputLogFilePath -Append
        
        # Active User Stories retrieving from the query : https://dev.azure.com/datosmic/datosmic-project1/_queries/query/f8266fd1-fdf3-4b81-ac14-155c4ba116d2/

        $ActiveUserStories = az boards query --id 'f8266fd1-fdf3-4b81-ac14-155c4ba116d2' | ConvertFrom-Json # User stories query 
        
        Write-Output "$(Get-Date): Provided values: `nOrganization: $Organization `nProject: $Project `nADO Query used: f8266fd1-fdf3-4b81-ac14-155c4ba116d2" | Out-File -FilePath $OutputLogFilePath -Append
        
        # Month Name for the retrieving the user story based on title

        $FullMonthName = (Get-UICulture).DateTimeFormat.GetMonthName((Get-Date).Month) + ' ' + ((Get-Date).Year)

        Write-Output "$(Get-Date): Generated Month Name for  User Story Title: $FullMonthName" | Out-File -FilePath $OutputLogFilePath -Append

        # Generating Week start date and Week end date for task title in format [ Monday's date in MM/DD - Sunday's date in MM/DD] Efforts Tracker - TeamName
        # Determine nearest Monday of the week
        $weekStartDate = Get-Date
        Write-Output "$(Get-Date): Determining the nearest Monday of the week... `n"
        while ($weekStartDate.DayOfWeek -ne 'Monday') { $weekStartDate = $weekStartDate.AddDays(-1) } 
        $weekStartDate = $weekStartDate.ToString('MM/dd')

        # Determine next Sunday of the week
        $weekEndDate = Get-Date
        Write-Output "$(Get-Date): Determining the nearest Sunday of the week... `n"
        while ($weekEndDate.DayOfWeek -ne 'Sunday') { $weekEndDate = $weekEndDate.AddDays(1) } 
        $weekEndDate = $weekEndDate.ToString('MM/dd')

        Write-Output "$(Get-Date): Generated Week Start Date and Week End Date for Task Title: `nWeek Start Date: $weekStartDate `nWeek End Date: $weekEndDate" | Out-File -FilePath $OutputLogFilePath -Append
   
        $Teams = @{
   
            Sales            = @{
                TeamName      = 'Sales & Marketing'
                TeamLeadAlias = 'MiriamG'
                TeamMembers   = @('MiriamG', 'MeganB', 'AlexW', 'IsaiahL', 'LynneR', 'AdeleV')
            }
        
            Operations        = @{
                TeamName      = 'Operations'
                TeamLeadAlias = 'NestorW'
                TeamMembers   = @('NestorW', 'JoniS', 'PradeepG', 'DiegoS')
            }
        
            Manufacturing     = @{
                TeamName      = 'Manufacturing'
                TeamLeadAlias = 'LeeG'
                TeamMembers   = @('LeeG', 'HenriettaM', 'GradyA', 'LidiaH', 'JohannaL')
            }

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
        
        $Tasks = @()
        
        $formattedTasks = @()

        foreach ($Team in $Teams.Keys) {

            $TeamMember = $Teams.$Team.TeamMembers
        
            $TeamMember | ForEach-Object {
        
                $UserStoryTitle = "[$FullMonthName] $($Teams.$Team.TeamName) Efforts Tracker"
                
                $ParentId = ($ActiveUserStories | Where-Object { $_.fields.'System.Title' -eq $UserStoryTitle }).Id
        
                $TaskTitle = "[$weekStartDate - $weekEndDate] Efforts Tracker - $($Teams.$Team.TeamName)" 
                
                Write-Output "$(Get-Date): Passing Values to create Task - [New-ADOWorkItem function]: `nOrganization: $($organization) `nProject: $($Project) `nTeamName: $($Teams.$Team.TeamName) `nTeamLeadAlias: $($Teams.$Team.TeamLeadAlias) `nTeamMemberName: $_ `nUserStoryTitle: $UserStoryTitle `nParentId: $ParentId `nTaskTitle: $TaskTitle `nTags: $($Teams.$Team.TeamName) `n" | Out-File -FilePath $OutputLogFilePath -Append
                
                $Tasks += New-ADOWorkItem -organization $organization -Project $project -Title $TaskTitle -WorkItemType $WorkItemType -AssignedTo $_ -ParentID $ParentId -Tags $($Teams.$Team.TeamName)
        
            }
        }

        # Sending Mail to the Team with created work items

        $Tasks | ForEach-Object {

            $f = $($_.fields)
            $formattedTask = @(
                [PSCustomObject]@{
                    'Work Item ID' = "<a href='https://dev.azure.com/datosmic/datosmic-project1/_workitems/edit/$($_.Id)'>$($_.Id)</a>"
                    'Title'        = $($f.'System.Title');
                    'Assigned To'  = $($f.'System.AssignedTo'.'displayName');
                    'Team'         = $($f.'System.Title'.Split('-')[-1].trim());
                    'Project'      = $($f.'System.TeamProject');
                    'Iteration'    = $($f.'System.IterationPath');
                    'State'        = $($f.'System.State');
                    'Type'         = $($f.'System.WorkItemType');
                    'Tags'         = $($f.'System.Tags');
                    'Parent ID'    = "<a href='https://dev.azure.com/datosmic/datosmic-project1/_workitems/edit/$($f.'System.Parent')'>$($f.'System.Parent')</a>"
                              
                }
            )
        
            $formattedTasks += $formattedTask
        }

# Table styling for mail
$styling = @"
<style>
th, td {
    padding: 12px;
    }
    table, th, td {
    border: 1px solid black;
    border-collapse: collapse;
    text-align:center;
}
th {
    background-color: #008080;
    color:white;
}
li {
    color:red;
}
</style>
"@
        
# Mail Content

$IntroStatement = @"
<p style='margin:0in;font-size:15px;font-family:"Calibri",sans-serif;color:#470FF4;'>Hello Team,</p>
<br/>
<ol style="margin-bottom:0in;margin-top:0in;" start="1" type="1">
    <li style='margin-top:0in;margin-right:0in;margin-bottom:0in;margin-left:0in;font-size:15px;font-family:"Calibri",sans-serif;'>Make sure to <strong>CLOSE</strong> your previous tasks with the <strong>Completed Work</strong> field filled with your complete efforts of the week.</li>
    <li style='margin-top:0in;margin-right:0in;margin-bottom:0in;margin-left:0in;font-size:15px;font-family:"Calibri",sans-serif;'><strong>DO NOT</strong> close User Stories until month end. New User Stories will be created every month and will be moved to respective iteration (1 month).</li>
    <li style='margin-top:0in;margin-right:0in;margin-bottom:0in;margin-left:0in;font-size:15px;font-family:"Calibri",sans-serif;'>New tasks will be created every week. Please change your respective tasks state from <strong>New -&gt; Active</strong> and move to current iteration.</li>
</ol>
<br/>
"@

$TaskWeekStatement = "<p style='color:#470FF4;'>Please find the list of tasks created for the week: $($weekStartDate) - $($weekEndDate) :</p><br/>"

$Note = "<br/><br/><p style='color:#470FF4;'><i>Note: This is an auto generated Email using PowerShell sent by Ravi Kiran Srikantam. Please reach out to <a href='mailto:scvslsravikiran@tm8h1.onmicrosoft.com'>Ravi Kiran Srikantam</a> for any queries/clarifications. Have a great day!</i></p>"

$Content = $formattedTasks | ConvertTo-Html -As Table -PreContent "$IntroStatement $TaskWeekStatement" -Head $styling -PostContent $Note

$HTMLMailContent = $Content -replace ('&lt;','<') -replace('&gt;','>') -replace ('&#39;','') 

Write-Verbose "Generated HTML Mail Content: `n`n$HTMLMailContent"

# Create an instance of Outlook.Application COM object
$outlook = New-Object -ComObject Outlook.Application

# Get the MAPI namespace
$namespace = $outlook.GetNamespace('MAPI')

# Construct email item object
$mailItem = $outlook.CreateItem(0)
$mailItem.Subject = "[DatOsmic Team] Efforts Tasks for the week: $($weekStartDate) - $($weekEndDate)"

$mailItem.BodyFormat = 2 # HTML Format
$mailItem.HTMLBody = "$($HTMLMailContent)"
$mailItem.To = "scvslsravikiran@tm8h1.onmicrosoft.com"
$mailItem.Cc = "PattiF@tm8h1.onmicrosoft.com"
$mailItem.Sensitivity = 2

$mailItem.Display()
$mailItem.Save()
# $mailItem.Send()

# Release the COM objects
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($mailItem) | Out-Null
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($namespace) | Out-Null
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($outlook) | Out-Null


    }
    
    end {
        return $Tasks   
    }
}











