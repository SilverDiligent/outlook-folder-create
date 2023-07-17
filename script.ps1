#load config data from json file
$configData = get-content -path ./config.json | ConvertFrom-Json

$clientId = $configData.client_id
$clientSecret = $configData.client_secret
$tenantId = $configData.tenant_id

# Import necessary module
Import-Module MSAL.PS

# Define client ID, tenant ID, and client secret obtained from Azure AD App registration
$authority = "https://login.microsoftonline.com/$tenantId"
$resourceUrl = "https://graph.microsoft.com"


# Convert client secret to SecureString
$secretSecureString = ConvertTo-SecureString -String $clientSecret -AsPlainText -Force

# Request access token
$tokenRequest = Get-MsalToken -ClientId $clientId -TenantId $tenantId -ClientSecret $secretSecureString -Authority $authority -Scope "$resourceUrl/.default"

# Define header with access token
$headers = @{
  "Authorization" = "Bearer $($tokenRequest.AccessToken)"
  "Content-Type"  = "application/json"
}

# Define the new folder name
$folderName = "FolderToCreate"

# Import user mailboxes from a CSV file
$mailboxes = Import-Csv -Path ./accounts.csv

# Loop through each mailbox and create a new folder
foreach ($mailbox in $mailboxes) {
  # Define the API endpoint and body
  # $uri = "$resourceUrl/v1.0/users/$($mailbox.Email)/mailFolders/Inbox/childFolders"
  $uri = "$resourceUrl/v1.0/users/$($mailbox.Email)/mailFolders/"
  $body = @{
    "displayName" = $folderName
  } | ConvertTo-Json

  # Make the API request
  $response = Invoke-RestMethod -Uri $uri -Method Post -Body $body -Headers $headers

  # Check the result
  if ($response) {
    Write-Host "Successfully created folder '$folderName' in the Inbox of $($mailbox.Email)"
  }
  else {
    Write-Host "Failed to create folder in $($mailbox.Email)"
  }
}
