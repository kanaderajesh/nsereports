Import-Module BitsTransfer

Function Send-Report {
[cmdletbinding()]
Param (
[string]$EmailTo, 
[String]$Subject,
[String]$Body
   ) 
# End of Parameters
Process {
    $EmailFrom = $Env:EmailFrom
    $EmailTo = $Env:EmailTo 
    $Subject = "Bulk Trading report "
    $SMTPServer = "smtp.gmail.com" 
    $Message = New-Object System.Net.Mail.MailMessage $EmailFrom, $EmailTo
    $Message.IsBodyHtml = $false
    $Message.Subject = $Subject
    $Message.body = $Body
    $SMTPClient = New-Object Net.Mail.SmtpClient($SmtpServer, 587) 
    $SMTPClient.EnableSsl = $true 
    $SMTPClient.Credentials = New-Object System.Net.NetworkCredential($Env:EmailFrom, $Env:Password); 
    $SMTPClient.Send($Message)
    
    } # End of Process
}
Function Get-Report {
[cmdletbinding()]
Param (
[string]$Url, 
[String]$Output
   ) 
# End of Parameters
Process {
    $start_time = Get-Date
    
    Start-BitsTransfer -Source $Url -Destination $Output
    #OR
    Start-BitsTransfer -Source $Url -Destination $Output -Asynchronous

    Write-Output "Time taken: $((Get-Date).Subtract($start_time).Seconds) second(s)" 

    } # End of Process
}

$filename = Get-Date -Format "ddMMyyyy"

$bulkurl = "https://www.nseindia.com/content/equities/bulk.csv"
$bulkoutput = "C:\Users\family\Downloads\nsereports\$filename" + "_bulk.csv"
if (-not ( Test-Path $bulkoutput) ){
    Get-Report -Url $bulkurl -Output $bulkoutput
}else{
    Write-Host ("File {0} already downloaded" -f $bulkoutput )
}

$blockurl = "https://www.nseindia.com/content/equities/block.csv"
$blockoutput = "C:\Users\family\Downloads\nsereports\$filename" + "_block.csv"
if (-not (Test-Path $blockoutput )){
    Get-Report -Url $blockurl -Output $blockoutput
}else{
     Write-Host ("File {0} already downloaded" -f $blockoutput )
}

$bulkbuy = Import-Csv -Path $bulkoutput -Header "Date","Symbol","Security Name","Client Name","Buy/Sell","Quantity Traded","Trade Price / Wght. Avg. Price","Remarks" | Select-Object -Skip 1
$Body = $bulkbuy| Sort-Object -Property "Quantity Traded" | Format-Table Symbol, Buy/Sell, "Quantity Traded" | Out-String
Write-Host $Body
Send-Report -EmailTo $EmailTo -Subject $Subject -Body $Body

$deravatives = "MTO_01032018.DAT"

$derivativesurl = "https://www.nseindia.com/archives/equities/mto/$deravatives"
$derivativesoutput = "C:\Users\family\Downloads\nsereports\$deravatives"
if (-not (Test-Path $derivativesoutput )){
    Get-Report -Url $derivativesurl -Output $derivativesoutput
}else{
     Write-Host ("File {0} already downloaded" -f $derivativesoutput )
}
