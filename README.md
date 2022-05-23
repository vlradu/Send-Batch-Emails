# Aspire-Send-Batch-Emails
This is a short PowerShell script used to automate the sending of buddy pairing emails for new waves of Microsoft Aspires.  
Can be used as a general example for sending batches of emails in any scenario.

```PowerShell
<#
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
