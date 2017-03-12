# Configuration
## E-mail
$smtpServer = "127.0.0.1"
$smptUser = "username"
$smtpPassword = "password"
$fromAddress = "oumonitor@company.com"
$recipientAddresses = "address1@company.com",
                      "address2@company.com"
$emailSubject = "OU List"
$emailBody = "Hi There!,

Please find attached recent user list in specified OUs.

Cheers,
Your System."

## Active Directory
$ouList  = "OU=MyOU,dc=domain,dc=company,dc=com",
           "OU=AnotherOU,dc=domain,dc=company,dc=com"
$queryFilter = "*"
$requiredUserProperties = "mail","Department"
$outputFields = "GivenName","Surname","mail","Department"
$outputFileName = "c:\temp\ad-users.csv"

# Load required modules
$checkCmds = Get-Command -Module ActiveDirectory
Import-Module ActiveDirectory
$userList = New-Object System.Collections.ArrayList

# Pull users from every OU into array
foreach ($ouName in $ouList) {
  $userList += Get-ADUser -SearchBase $ouName -Filter $queryFilter -Properties $requiredUserProperties | Select $outputFields
}

# Write user list to a CSV file
$userList | Export-Csv -path $outputFileName -NoTypeInformation

# Compose e-mail with attachment
$attachedFile = new-object Net.Mail.Attachment($outputFileName)
$message = new-object Net.Mail.MailMessage
$smtp = new-object Net.Mail.SmtpClient($smtpServer)
$smtp.Credentials = New-Object System.Net.NetworkCredential( $smtpUser , $smtpPassword );
$message.From = $fromAddress
foreach ($recipient in $recipientAddresses) {
  $message.To.Add($recipient)
}
$message.Subject = $emailSubject
$message.Body = $emailBody
$message.Attachments.Add($attachedFile)

# Send e-mail and discard object
$smtp.Send($message)
$attachedFile.Dispose()

# Unload if it was not loaded before
if ($checkCmds -eq $null) { Remove-Module ActiveDirectory }
