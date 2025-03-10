#Changelog
# 2025-04-03: Initial version 1.0
# 2025-04-04: Added logging and error handling - version 1.1
# 2025-04-06: Added token handling - version 1.2
# 2025-04-07: Added token protection and removed clear text credentials - version 1.3
# 2025-04-10: Improved initial credential setup - version 1.3.1


# Set working directory
Set-Location -Path "C:\Script"

# Configuration - Fill these in
$username = "username@domain.nl"             # Your Brandweerrooster username (usually email address)

# Logging function
function Log-Message {
    param (
        [string]$Message,
        [string]$Level = "INFO"  # Can be INFO, ERROR, WARNING, etc.
    )
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logMessage = "[$timestamp] [$Level] $Message"
    Add-Content -Path ".\Brandweerrooster_log.txt" -Value $logMessage
    Write-Output $logMessage  # Optionally still show in console
}

Log-Message "Script started."

# Path to the password file
$passwordFile = ".\password.txt"

# Check if password file exists
if (-Not (Test-Path $passwordFile)) {
    Write-Host "No stored password found. Please enter your password."
    $securePassword = Read-Host "Enter password" -AsSecureString
    $securePassword | ConvertFrom-SecureString | Set-Content -Path $passwordFile
    Write-Host "Password saved securely."
    Log-Message "New password securely stored."
} else {
    Log-Message "Using stored secure password."
}

# Read and decrypt the stored password
$securePassword = Get-Content $passwordFile | ConvertTo-SecureString
$password = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto(
    [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($securePassword)
)

# API URL and information
$ApiBaseUrl = "https://www.brandweerrooster.nl/api/v2"
$MembershipID = "12345" # Your membership ID, retrieve from brandweerrooster URL
$tokenUrl = "https://www.brandweerrooster.nl/oauth/token"
$tokenFile = ".\brandweerrooster_token.json"  # Path where the encrypted token will be saved

# Function to save encrypted token (only access_token is encrypted)
function Save-EncryptedToken {
    param (
        [object]$TokenData
    )
    $plainToken = $TokenData.access_token

    # Encrypt just the token
    $secureString = ConvertTo-SecureString -String $plainToken -AsPlainText -Force
    $encryptedToken = ConvertFrom-SecureString -SecureString $secureString

    # Create a safe object with encrypted token and plain metadata
    $safeTokenData = [PSCustomObject]@{
        expires_at            = $TokenData.expires_at
        token_type             = $TokenData.token_type
        scope                  = $TokenData.scope
        access_token_encrypted = $encryptedToken
    }

    # Save to file
    $safeTokenData | ConvertTo-Json -Depth 10 | Set-Content -Path $tokenFile -Encoding UTF8
    Log-Message "Encrypted token saved to file."
}

# Function to load and decrypt token
function Load-EncryptedToken {
    if (Test-Path $tokenFile) {
        try {
            $safeTokenData = Get-Content -Raw -Path $tokenFile | ConvertFrom-Json

            # Decrypt just the access_token
            $secureString = ConvertTo-SecureString -String $safeTokenData.access_token_encrypted
            $plainToken = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto(
                [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($secureString)
            )

            # Rebuild the full token object with decrypted access_token
            $fullTokenData = [PSCustomObject]@{
                expires_at    = $safeTokenData.expires_at
                token_type    = $safeTokenData.token_type
                scope         = $safeTokenData.scope
                access_token  = $plainToken
            }

            return $fullTokenData
        } catch {
            Log-Message "Failed to read or decrypt saved token: $($_.Exception.Message)" -Level "WARNING"
            return $null
        }
    } else {
        return $null
    }
}

# Token management functions
function Get-NewToken {
    Log-Message "Requesting new access token..."
    $body = @{
        grant_type = "password"
        username   = $username
        password   = $password
    }

    try {
        $response = Invoke-RestMethod -Uri $tokenUrl -Method Post -Body $body
        $response | Add-Member -MemberType NoteProperty -Name expires_at -Value ((Get-Date).AddSeconds($response.expires_in))
        Save-EncryptedToken -TokenData $response
        Log-Message "New access token obtained and saved."
        return $response
    } catch {
        Log-Message "Failed to obtain new token: $($_.Exception.Message)" -Level "ERROR"
        exit
    }
}

function Get-ValidToken {
    $tokenData = Load-EncryptedToken
    if ($tokenData -and ([datetime]$tokenData.expires_at -gt (Get-Date))) {
        Log-Message "Using existing valid encrypted token."
        return $tokenData
    } else {
        Log-Message "No valid saved token found (or expired), requesting new token."
        return Get-NewToken
    }
}

# Get a valid token (either existing or new)
$tokenData = Get-ValidToken
$accessToken = $tokenData.access_token

# Set up headers
$headers = @{
    Authorization = "Bearer $accessToken"
}

# Date range processing
$midnightToday = Get-Date -Hour 0 -Minute 0 -Second 0 -Millisecond 0 -Format "yyyy-MM-ddTHH:mm:ssK"
# URL encode (replace : and +)
$encodedMidnightToday = $midnightToday -replace ":", "%3A" -replace "\+", "%2B"
# Midnight 30 days from now
$midnight30Days = (Get-Date).AddDays(30).Date.ToString("yyyy-MM-ddTHH:mm:ssK")
# URL encode
$encodedMidnight30Days = $midnight30Days -replace ":", "%3A" -replace "\+", "%2B"

# Fetch raw availability data
$Brandweerrooster_raw_output = ".\Brandweerrooster_raw_output.txt"  # Path where the json file will be saved
try {
    $availabilityData = Invoke-RestMethod -Uri "$ApiBaseUrl/memberships/$MembershipID/combined_schedule?start_time=$encodedMidnightToday&end_time=$encodedMidnight30Days" -Headers $headers -Method Get | ConvertTo-Json -Depth 10
    Write-Output "Successfully retrieved availability data."
    out-file -FilePath $Brandweerrooster_raw_output -InputObject $availabilityData -Encoding UTF8
} catch {
    Log-Message "Failed to retrieve availability data: $($_.Exception.Message)" -Level "ERROR"
    exit
}

# Remove block where availability is set to false from the json file
$Brandweerrooster_schedule = ".\Brandweerrooster_schedule.txt"
# Read raw JSON file
try {
    $jsonData = Get-Content -Raw -Path $Brandweerrooster_raw_output | ConvertFrom-Json
    $jsonData.intervals = @($jsonData.intervals | Where-Object { $_.available -eq $true })
    $jsonData.intervals | Out-File -FilePath $Brandweerrooster_schedule -Encoding UTF8
    Log-Message "Filtered schedule saved to $Brandweerrooster_schedule"
} catch {
    Log-Message "Failed to process and save schedule data: $($_.Exception.Message)" -Level "ERROR"
    exit
}

#Convert JSON to ICS
# Define file paths
$inputFilePath = $Brandweerrooster_schedule
$outputICSPath = ".\brandweercalendar.ics"  # Adjust to your webserver path

try {
    $fileContent = Get-Content -Path $inputFilePath
    Log-Message "Read filtered schedule file."

    $calendarItems = @()
    $currentItem = @{}

    foreach ($line in $fileContent) {
        if ([string]::IsNullOrWhiteSpace($line)) {
            if ($currentItem.Count -gt 0) {
                $calendarItems += [PSCustomObject]$currentItem
                $currentItem = @{}
            }
        }
        else {
            $parts = $line -split ":", 2
            $key = $parts[0].Trim()
            $value = $parts[1].Trim()

            if ($value -match '^{.*}$') {
                $value = $value -replace "[{}]", ""
                $value = $value -split "," | ForEach-Object { $_.Trim() }
            }

            $currentItem[$key] = $value
        }
    }

    if ($currentItem.Count -gt 0) {
        $calendarItems += [PSCustomObject]$currentItem
    }

    Log-Message "Parsed $($calendarItems.Count) calendar items."

    $icsContent = @()
    $icsContent += "BEGIN:VCALENDAR"
    $icsContent += "VERSION:2.0"
    $icsContent += "PRODID:-//MyOrg//Availability Calendar//EN"
    $icsContent += "CALSCALE:GREGORIAN"

    foreach ($item in $calendarItems) {
        $start = [DateTimeOffset]::Parse($item.start_time).UtcDateTime.ToString("yyyyMMddTHHmmssZ")
        $end = [DateTimeOffset]::Parse($item.end_time).UtcDateTime.ToString("yyyyMMddTHHmmssZ")
        $uid = [guid]::NewGuid().ToString()

        $skills = if ($item.skill_ids -is [array]) { $item.skill_ids -join ", " } else { $item.skill_ids }

        $icsContent += "BEGIN:VEVENT"
        $icsContent += "UID:$uid"
        $icsContent += "DTSTAMP:$(Get-Date -Format 'yyyyMMddTHHmmssZ')"
        $icsContent += "DTSTART:$start"
        $icsContent += "DTEND:$end"
        $icsContent += "SUMMARY:Brandweer dienst"
        $icsContent += "DESCRIPTION:Brandweer dienst. Skills: $skills"
        $icsContent += "TRANSP:TRANSPARENT"
        $icsContent += "END:VEVENT"
    }

    $icsContent += "END:VCALENDAR"

    $icsContent -join "`r`n" | Set-Content -Path $outputICSPath -Encoding UTF8
    Log-Message "ICS file written to $outputICSPath"
} catch {
    Log-Message "Failed to convert schedule to ICS file: $($_.Exception.Message)" -Level "ERROR"
    exit
}

Log-Message "Script completed successfully."