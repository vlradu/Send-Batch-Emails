<#
.NOTES
    AspireAutomatedEmails.ps1

    Written by Vlad Radu (vlradu@microsoft.com), 2022.

    Requires the Microsoft.Graph.Authentication and Microsoft.Graph.Users PowerShell modules to run.
    Assumes the csv contains 4 columns named NewHireEmail,BuddyEmail,NewHireName,BuddyName. If names
    are different, this can be configured when calling the Send-AspireEmail function (e.g., change
    $Pair.NewHireName to $Pair.NewAspireName if the column is named NewAspireName).

    To be used internally to automate the sending of emails to new waves of Aspires. 
    Use at your own risk.  No warranties are given.
.SYNOPSIS
    Basic script that uses Graph API PowerShell SDK to send batches of emails.
    
.DESCRIPTION
    This script authenticate you in the browser via Modern Authentication, then 
    will take a .csv file and loop through all lines to send emails to the new 
    aspires and their buddies. Can configure the HTML body via the $MailContent variable.
    The script stores running logs in the log/ folder and any sending errors in new CSV files 
    in the Failed/ folder. These CSVs can then be used to re-send any failed emails, they use the
    same format as the original CSV.

.PARAMETER CsvFile
    This is the full or relative path to the CSV file which contains all pairs of new Aspires and 
    their buddies. Mandatory parameter. Can also be used to point to a CSV of failed recipients to 
    retry sending.

.PARAMETER SendFrom
    This is an optional parameter, you don't necessarily need to send it if you authenticate to the 
    mailbox you'll send emails from. You need to use it if you log in to one mailbox and wish to 
    send from another mailbox via SendAs/SendOnBehalf permissions (e.g from a shared mailbox).

.PARAMETER HtmlFile
    This is an optional parameter, you can set it to the path of a .html template is saved, to be used
    instead of setting the $MailContent variable directly in this script.
    NOTE: In order to personalize the email, $AspireName needs to be replaced with {0} and 
    $BuddyName with {1} in the text of the HTML file, they will point to the Aspires' and buddies' names.
.EXAMPLE
    .\AspireAutomatedEmails.ps1 -CsvFile ".\ExampleList.csv"

.EXAMPLE
    .\AspireAutomatedEmails.ps1 -CsvFile ".\ExampleList.csv" -SendFrom user@example.com 

.EXAMPLE
    .\AspireAutomatedEmails.ps1 -CsvFile ".\ExampleList.csv" -HtmlFile ".\ExampleHtml.html"    

.EXAMPLE
    .\AspireAutomatedEmails.ps1 -CsvFile ".\ExampleList.csv" -SendFrom user@example.com -HtmlFile ".\ExampleHtml.html"  
#>

param(
    [Parameter(Mandatory = $true,
        HelpMessage = "Path of CSV file containing Aspire list",
        Position = 0)]
    [string] $CsvFile,
    [Parameter(Mandatory = $false,
        HelpMessage = "Set this property to the mailbox you're sending from. `
        Optional, only needed if sending from different mailbox than your own through SendAs/SendOnBehalf permissions",
        Position = 1)]
    [string] $SendFrom,
    [Parameter(Mandatory = $false,
        HelpMessage = "Optional. Set this to the path of a HTML template you'll be using.",
        Position = 2)]
    [string] $HtmlFile
)

#Prerequisites to run this:
#Install-Module Microsoft.Graph.Authentication
#Install-Module Microsoft.Graph.Users


Import-Module Microsoft.Graph.Users
Connect-MgGraph -Scopes "Mail.Send","Mail.Send.Shared"

function Send-AspireEmail
{
    param(
        $AspireEmail,
        $BuddyEmail,
        $AspireName,
        $BuddyName,
        $SenderEmail,
        $HtmlFile = $null
    )
    try
    {
        if($HtmlFile -ne $null)
        {
            $MailContent = Get-Content $HtmlFile -Raw | Out-String #| ConvertTo-HTML -Fragment | Out-Null
            #$MailContent = "$MailContent"
            $MailContent = $MailContent -f $AspireName, $BuddyName
        }
        else
        {
            #Don't forget to include the $AspireName and $BuddyName variables in the mail body to personalize it a bit.
            $MailContent = @"
            <table style="width: 100%; border-collapse: collapse; background-color: #0000ff;" border="1">
            <tbody>
            <tr>
            <td style="width: 100%; text-align: center;"><span style="color: #ffffff;">Aspire Buddy Program</span></td>
            </tr>
            <tr style="text-align: center;">
            <td style="width: 100%;">
            <p><span style="color: #ffffff;">Hi there $AspireName and $BuddyName,</span></p>
            <p><span style="color: #ffffff;">We've paired you up. K thx bye.</span></p>
            </td>
            </tr>
            </tbody>
            </table>
"@  #Leave this unindented.
        }
    }
    catch
    {
        throw
    }


    $MailBody = @{
        Message = @{
            Subject = "Test Subject";
            Body = @{
                Content = $MailContent; 
                ContentType = "HTML"
                }; 
            ToRecipients = @(
                @{
                emailAddress = @{
                    address = $AspireEmail
                    }
                };
                @{
                    emailAddress = @{
                        address = $BuddyEmail
                        }
                }
            )
            <#BccRecipients = @(    #Uncomment this block if you want to add any Bcc recipients
                @{ 
                emailAddress = @{
                    address = "recipient@example.com"
                    }
                }                          
            )#>
        }
        savetoSentItems = "true"
    }
    $Response = Send-MgUserMail -UserId $SenderEmail -BodyParameter $MailBody -PassThru
    if($Response -ne $True)
    {
        throw 
    }
}

$AspiresList = Import-CSV -Path $CsvFile

if($PSBoundParameters.ContainsKey('SendFrom'))
{
    $SenderAddress = $SendFrom
}
else
{
    $SenderAddress = (Get-MgContext).Account
}

if(-not($PSBoundParameters.ContainsKey('HtmlFile')))
{
    $HtmlFile = $null
}


if(-not(Test-Path ".\Log\"))
    {
        New-Item ".\Log" -ItemType Directory
    }
$CurrentDate = Get-Date -Format "yyMMddHHmmssffff"
Start-Transcript -Path ".\Log\$CurrentDate.txt"
$CountFailed = 0
foreach ($Pair in $AspiresList) {
    Start-Sleep 2 #Waiting 2 seconds between each email because EXO has a 30 mails/minute rate limit (60/2=30)
    try{
        Write-Host "Sending email to $Pair"
        Send-AspireEmail -AspireEmail $Pair.NewHireEmail -BuddyEmail $Pair.BuddyEmail `
        -AspireName $Pair.NewHireName -BuddyName $Pair.BuddyName -SenderEmail $SenderAddress -HtmlFile $HtmlFile
        Write-Host "Successful." -ForegroundColor Green
    }
    catch
    {
        $CountFailed += 1
        Write-Host "Failed to send to $Pair. `nAdding pair to Failed.csv" -ForegroundColor Red
        Write-Host "Error: $(Get-Error -Newest 1)" -ForegroundColor Red
        $Pair | Export-Csv "Failed.csv" -Append
    }
}

if(-not(Test-Path ".\Failed\"))
{
    New-Item ".\Failed" -ItemType Directory
}

if(Test-Path ".\Failed.csv")
{
    $ErrorCsv = "$CurrentDate.csv"
    Move-Item ".\Failed.csv" ".\Failed\$ErrorCsv"
    Write-Host "Saved $CountFailed failed email(s) to $.\Failed\$ErrorCsv" -ForegroundColor Magenta
}
else
{
    Write-Host "All emails sent successfully." -ForegroundColor Cyan
}

Stop-Transcript
Disconnect-MgGraph