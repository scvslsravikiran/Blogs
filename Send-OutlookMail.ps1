$process = Get-Process -ProcessName msedge | Select-Object Name, Id, ProcessName, CPU

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
    color:#470FF4;
}
</style>
"@
        
# Mail Content

$IntroStatement = @"
<p style='margin:0in;font-size:15px;font-family:"Calibri",sans-serif;color:#470FF4;'>Hello Team,</p>
<br/>
<ol style="margin-bottom:0in;margin-top:0in;" start="1" type="1">
    <li style='margin-top:0in;margin-right:0in;margin-bottom:0in;margin-left:0in;font-size:15px;font-family:"Calibri",sans-serif;'>Please find  <strong>MS EDGE</strong> instances running in the machine currently.</li>
</ol>
<br/>
"@

$Note = "<br/><br/><p style='color:#470FF4;'><i>Note: This is an auto generated Email using PowerShell sent by Ravi Kiran Srikantam. Please reach out to <a href='mailto:scvslsravikiran@tm8h1.onmicrosoft.com'>Ravi Kiran Srikantam</a> for any queries/clarifications. Have a great day!</i></p>"

$Content = $process | ConvertTo-Html -As Table -PreContent "$IntroStatement" -Head $styling -PostContent $Note

$HTMLMailContent = $Content -replace ('&lt;','<') -replace('&gt;','>') -replace ('&#39;','') 

Write-Verbose "Generated HTML Mail Content: `n`n$HTMLMailContent"

# Create an instance of Outlook.Application COM object
$outlook = New-Object -ComObject Outlook.Application

# Get the MAPI namespace
$namespace = $outlook.GetNamespace('MAPI')

# Construct email item object
$mailItem = $outlook.CreateItem(0)
$mailItem.Subject = "[DatOsmic Team] Process of MS Edge - $(Get-Date)"

$mailItem.BodyFormat = 2 # HTML Format
$mailItem.HTMLBody = "$($HTMLMailContent)"
$mailItem.To = "scvslsravikiran@tm8h1.onmicrosoft.com"
$mailItem.Cc = "scvslsravikiran@tm8h1.onmicrosoft.com"
$mailItem.Sensitivity = 2

$mailItem.Display()
$mailItem.Save()
# $mailItem.Send()

# Release the COM objects
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($mailItem) | Out-Null
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($namespace) | Out-Null
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($outlook) | Out-Null
