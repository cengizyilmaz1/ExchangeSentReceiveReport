<#
=============================================================================================
Name = Cengiz YILMAZ
Microsoft Certified Trainer (MCT)
Date = 4.03.2023
www.cengizyilmaz.net
www.cozumpark.com/author/cengizyilmaz
============================================================================================
#>

$days = Read-Host "Enter the number of days to generate the report for"
$database = Read-Host "Enter the name of the Exchange database to scan. Press Enter to scan all mailboxes."

$properties = "Name", "DisplayName", "ItemCount", "LastLogonTime"

if ($database) {

    $mailboxes = Get-Mailbox -Database $database -ResultSize Unlimited -ErrorAction Stop | Select-Object $properties

} else {

    $mailboxes = Get-Mailbox -ResultSize Unlimited -ErrorAction Stop | Select-Object $properties

}

 
$tableRows = "<tr><th style='color: white;'>Name</th><th>EmailAddress</th><th style='color: white;'>Received Count</th><th style='color: white;'>Sent Count</th></tr>"

foreach ($mailbox in $mailboxes) {

    $email = $mailbox.Name

    $stats = Get-MailboxStatistics $email -ErrorAction SilentlyContinue

    if ($stats) {

        $received = (Get-MessageTrackingLog -Recipients $email -Start (Get-Date).AddDays(-$days) -ResultSize Unlimited -EventId Receive).Count

        $sent = (Get-MessageTrackingLog -Sender $email -Start (Get-Date).AddDays(-$days) -ResultSize Unlimited -EventId Send).Count

        $tableRows += "<tr><td>$($mailbox.DisplayName)</td><td>$($mailbox.Name)</td><td style='color:blue'>$($received)</td><td style='color:blue'>$($sent)</td></tr>"

        Write-Host "Processed $($email): $($received) received, $($sent) sent" -ForegroundColor Yellow

    }

}

$html = @"

<!DOCTYPE html>

<html>

<head>

    <style>

        table, th, td {

            border: 1px solid black;

            border-collapse: collapse;

            font-family: Arial;

            font-size: 11pt;

            text-align: left;

            padding: 5px;

        }

        th {

            background-color: #337ab7;

            color: white;

        }

    </style>

</head>

<body>

    <h1 style='background-color: #337ab7; color: white;'>Exchange Mailbox Sent/Receive Report</h1>

    <table>

        $tableRows

    </table>

</body>

</html>

"@

$html | Out-File "C:\script\Report.html"

 

Write-Host "Report generated successfully." -ForegroundColor Green
