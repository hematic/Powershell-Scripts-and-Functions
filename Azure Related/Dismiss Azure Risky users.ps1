function Invoke-DecodeCertificate {
    param (
        [Parameter(Mandatory=$true)]
        [string]$inputFile,
  
        [Parameter(Mandatory=$true)]
        [string]$outputFile
    )
  
    try {
        # Attempt to decode the certificate
        $result = & certutil.exe -decode $inputFile $outputFile 2>&1
  
        # Check for errors in the result
        if ($LASTEXITCODE -ne 0) {
            throw $result
        }
  
        # If successful, display a success message
        Write-Host "Successfully decoded and exported the certificate."
    }
    catch {
        Write-Error "An error occurred while decoding: $_"
        exit 1
    }
  }
  
  # Authenticate and acquire a token
  $client_id = $ENV:ARM_CLIENT_ID
  $tenant_id = $ENV:ARM_TENANT_ID
  $cert_file_content = $ENV:ARM_CLIENT_CERTIFICATE_PATH
  $cert_file_path = 'c:\temp\cert.pfx'
  $cert_password = $ENV:ARM_CLIENT_CERTIFICATE_PASSWORD
  
  Invoke-DecodeCertificate -inputFile $cert_file_content -outputFile $cert_file_path
  
  $secure_password = ConvertTo-SecureString -String $cert_password -AsPlainText -Force
  $import = Import-PfxCertificate -Password $secure_password -FilePath $cert_file_path -CertStoreLocation Cert:\CurrentUser\My
  $import
  Connect-MgGraph -ClientId $client_id -TenantId $tenant_id -CertificateThumbprint $import.thumbprint -NoWelcome
  
  Try{
    # Fetch all risky users
    $riskyUsersUrl = "https://graph.microsoft.com/beta/identityProtection/riskyUsers"
    $riskyUsers = Invoke-MgGraphRequest -Method GET -Uri $riskyUsersUrl -ErrorAction Stop | select -ExpandProperty value
    $riskyUsers | select -Property userDisplayName, riskstate, risklevel | FT
  }
  Catch{
    Write-Error "Problem gathering risky users: $_"
    exit 1
  }
  
  Try{
    foreach ($user in $riskyUsers){
        # Dismissing the risk for a user
        $dismissUrl = "https://graph.microsoft.com/beta/riskyUsers/dismiss"
        $body = @{
            userIds = @("$($user.id)")
        } | ConvertTo-Json
  
        Invoke-MgGraphRequest -Method POST -Uri $dismissUrl -Headers @{ "Content-Type" = "application/json" } -Body $body -ErrorAction Stop
        Write-Host "Dismissed any risks for user:" $user.userPrincipalName
    }
  }
  Catch{
    Write-Error "Problem dismissing risky users: $_"
    exit 1
  }
  # Clean up and disconnect from the session
  Disconnect-MgGraph
  