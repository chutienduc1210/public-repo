# This script is used to created new mail-user in Exchange Online Environment from imported CSV file. The password then will be sent to enduser via email.
# This script also update general information about users and require change password at next logon

Set-ExecutionPolicy Unrestricted

# Fuction to send email using Powershell
function Send-Email {
    param (
        [string]$To,
        [string]$TargetUserName,
        [string]$Password,
        [string]$DisplayName,
        [PSCredential]$Cred          # Use PSCredential for username and password
    )
    $body = @"
<html>
<body>
    <p>Dear $DisplayName,</p>
    
    <p>Your Office 365 account has been set up. Below are your login credentials:</p>
    
    <p><strong>Username:</strong> $TargetUserName<br>
        <strong>Temporary Password:</strong> $Password</p>

    <p><strong>Instructions:</strong></p>
    <ol>
        <li>Go to <a href="https://portal.office.com" target="_blank">https://portal.office.com</a> to log in.</li>
        <li>Enter your username and the temporary password provided above.</li>
        <li>You will be prompted to change your password during your first login. Please choose a secure, memorable password.</li>
        <li>If you experience any issues, please contact IT support at support@yourdomain.com.</li>
    </ol>
    
    <p>Thank you, and welcome to the team!</p>
    <p>Best Regards,<br>Your IT Support Team</p>
</body>
</html>
"@
    # Send email
    Send-MailMessage -BodyAsHtml -Body $body -From "Do-not-reply@M365x61560265.OnMicrosoft.com" -SmtpServer smtp.office365.com -Port 587 -Subject "Your Office 365 Account Credentials" -To $To -UseSsl -Credential $Cred
}

function Generate-randomString {
    param (
        [int]$Length = 8,
        [string]$UpperCase = "ABCDEFGHJKMNPQRSTUVWXYZ",
        [string]$LowerCase = "abcdefghjkmnpqrstuvwxyz",
        [string]$Numbers = "123456789",
        [string]$SpecialChars = "~!@#$%^&*"
    )
    $AllCharacters = $UpperCase + $LowerCase + $Numbers + $SpecialChars
    $random = New-Object System.Random

    $selected_chars = @()

    $selected_chars += $UpperCase[$random.Next(0, $UpperCase.Length)]
    $selected_chars += $LowerCase[$random.Next(0, $LowerCase.Length)]
    $selected_chars += $Numbers[$random.Next(0, $Numbers.Length)]
    $selected_chars += $SpecialChars[$random.Next(0, $SpecialChars.Length)]
    $RemainingLength = $Length - $selected_chars.Count
    for ($i = 1; $i -le $RemainingLength; $i++) {
        $selected_chars += $AllCharacters[$random.Next(0, $AllCharacters.Length)]
    }
    $generated_passwd = ($selected_chars | Sort-Object { Get-Random }) -join ""
    return $generated_passwd
}

function Export-IfNotEmpty {
    param (
        [array]$DataArray,
        [string]$FilePath
    )

    if ($DataArray.Count -gt 0) {
        $DataArray | Export-Csv -Path $FilePath -NoTypeInformation -Force
    }
}


Import-Module -Name Microsoft.Graph.Users
Write-Host "Connecting to Microsoft Graph Users Powershell. Please enter your credentials with Global Administrator permission." -ForegroundColor Yellow
Connect-MgGraph -Scopes "User.ReadWrite.All, Directory.ReadWrite.All, Directory.AccessAsUser.All"

$users = Import-Csv .\users.csv
$existed_users = @()
$created_objects = @()
$failed_to_create = @()
$failed_to_send_email = @()

foreach ($user in $users) {
    $UserPrincipalName = $user.UserPrincipalName
    $MailNickname = $user.MailNickname
    $GivenName = $user.GivenName
    $Surname = $user.Surname
    $DisplayName = $user.DisplayName
    $MobilePhone = $user.MobilePhone
    $OfficeLocation = $user.OfficeLocation
    $JobTitle = $user.JobTitle
    $Department = $user.Department
    $CompanyName = $user.CompanyName
    $UsageLocation = $user.UsageLocation
    $Gmail = $user.Gmail

    $check = Get-MgUser -All | where {$_.UserPrincipalName -eq $UserPrincipalName}

    if ($check) {
        $existed_user += [PSCustomObject]@{
            UserPrincipalName = $UserPrincipalName
        }
    }
    else {
        $passwdprofile =  New-Object -TypeName Microsoft.Graph.PowerShell.Models.MicrosoftGraphPasswordProfile
        $passwdprofile.Password = Generate-randomString -Length 12
        $passwdprofile.ForceChangePasswordNextSignIn = $true
        $passwdprofile.ForceChangePasswordNextSignInWithMfa = $true
        try {
            $new_user = New-MgUser -AccountEnabled -UserPrincipalName $UserPrincipalName -MailNickname $MailNickname -GivenName $GivenName -Surname $Surname -DisplayName $DisplayName -MobilePhone $MobilePhone -OfficeLocation $OfficeLocation -JobTitle $JobTitle -Department $Department -CompanyName $CompanyName -UsageLocation $UsageLocation -PasswordProfile $passwdprofile -ErrorAction SilentlyContinue
            $created_objects += [PSCustomObject]@{
                UserPrincipalName = $UserPrincipalName
                Password = $passwdprofile.Password
                Gmail = $Gmail
                MailNickname = $MailNickname
                GivenName = $GivenName
                Surname = $Surname
                DisplayName = $DisplayName
                MobilePhone = $MobilePhone
                OfficeLocation = $OfficeLocation
                JobTitle = $JobTitle
                Department = $Department
                CompanyName = $CompanyName
                UsageLocation = $UsageLocation
            }
        }
        catch {
            $failed_to_create += [PSCustomObject]@{
                UserPrincipalName = $UserPrincipalName
            }
        }
        
    }
}

# Credential for send email
Write-Host "Please provide the credentials used to send passwords to end-users via email. The following requirements must be met to successfully send emails from Office 365:" -ForegroundColor Blue
Write-Host "- Basic Authentication must be allowed and enabled for the account used to send emails. For more information, see https://shorturl.at/z01rb on how to enable Basic Authentication for a specific account." -ForegroundColor Blue
Write-Host "- To use SMTP AUTH, you need to disable security defaults. For more information, see https://shorturl.at/bZ7Os" -ForegroundColor Blue
$send_email_cred = Get-Credential

foreach ($user in $created_objects) {
    $UserPrincipalName = $user.UserPrincipalName
    $Password = $user.Password
    $MailNickname = $user.MailNickname
    $GivenName = $user.GivenName
    $Surname = $user.Surname
    $DisplayName = $user.DisplayName
    $MobilePhone = $user.MobilePhone
    $OfficeLocation = $user.OfficeLocation
    $JobTitle = $user.JobTitle
    $Department = $user.Department
    $CompanyName = $user.CompanyName
    $UsageLocation = $user.UsageLocation
    $Gmail = $user.Gmail

    # Send password via email
    try {
        $send = Send-Email -To $Gmail -TargetUserName $UserPrincipalName -Password $Password -DisplayName $DisplayName -Cred $send_email_cred -ErrorAction SilentlyContinue
    }
    catch {
        $failed_to_send_email += [PSCustomObject]@{
            UserPrincipalName = $UserPrincipalName
        }
    }
}

Export-IfNotEmpty -DataArray $existed_users -FilePath .\Existed_users.csv
Export-IfNotEmpty -DataArray $created_objects -FilePath .\Created_users.csv
Export-IfNotEmpty -DataArray $failed_to_create -FilePath .\Failed_users_creation.csv
Export-IfNotEmpty -DataArray $failed_to_send_email -FilePath .\Failed_to_send_email.csv
