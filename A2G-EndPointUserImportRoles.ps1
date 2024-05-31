try {
    do {
        # Set the URL and parameters for obtaining access token
        $urlToken = "https://cob-uat.infocorpnow.com:22500/identity/connect/token"
        $headersToken = @{
            "Content-Type" = "application/x-www-form-urlencoded"
        }
        $bodyToken = @{
            client_id     = "cob_usermanager"
            client_secret = "xxxx"
            grant_type    = "client_credentials"
            scope         = "external.neworder.v2"
        }

        # Execute the request to obtain access token
        $responseToken = Invoke-RestMethod -Uri $urlToken -Method Post -Headers $headersToken -Body $bodyToken

        # Extract the access_token
        $accessToken = $responseToken.access_token

        # Set headers for subsequent API requests
        $headers = @{
            "Accept"        = "*/*"
            "Content-Type"  = "application/json"
            "Authorization" = "Bearer $accessToken"
        }
        $headers.Add("User-Agent", "Thunder Client (https://www.thunderclient.com)")

        # Prompt user for input
        $email = Read-Host -Prompt 'Enter email'
        $roleID = Read-Host -Prompt 'Enter role ID'
        $locationInput = Read-Host -Prompt 'Enter location IDs, separated by commas'

        # Convert locationIDs to an array
        $locationIDs = $locationInput -split ',' | ForEach-Object { $_.Trim() }

        # If there's only one location or the input is empty, consider it as a single-location array
        if ($locationIDs.Count -eq 1 -or $locationInput -eq "") {
            $locationIDs = @($locationInput.Trim())
        }

        # Extract employeeId from the email (use the whole part before the email domain)
        $emailParts = $email -split '@'
        $employeeId = $emailParts[0].Trim()

        # Extract parts of the email address
        $names = $emailParts[0] -split '\.'

        # Extract first name and last name, ignoring the middle name
        $firstName = $names[0].Trim()
        $lastName = $names[-1].Trim()

        # Initialize array to store responses
        $responses = @()

        $reqUrl = 'https://brampton-pay-uat.infocorpnow.com:22500/Order/api/v1/Employee'

        # Display information for user verification
        Write-Host "`nVerify the following information:`n"
        Write-Host "  Email: $email"
        Write-Host "  Employee ID: $employeeId"
        Write-Host "  lastName  : $lastName"
        Write-Host "  firstName : $firstName"
        Write-Host "  Role ID: $roleID"
        Write-Host "  Location IDs: $($locationIDs -join ', ')"
        Write-Host "  Date of Hire: 1901-01-01T00:00:00"  # Hardcoded date, adjust as needed
        Write-Host "  Date Left: 1901-01-01T00:00:00"  # Hardcoded date, adjust as needed
        Write-Host "dateOfBirth: 1901-01-01T00:00:00"
        Write-Host "reviewDate:  1901-01-01T00:00:00"

        # Prompt user for confirmation
        $confirmation = Read-Host "`nAre the details correct? (Y/N)"
        if ($confirmation -eq 'Y' -or $confirmation -eq 'y') {

            # Define the JSON body using user input
            $body = @{
                "employeeId"           = $employeeId
                "email"                = $email
                "lastName"             = $lastName
                "firstName"            = $firstName
                "middleName"           = ""
                "addressLine1"         = ""
                "addressLine2"         = ""
                "city"                 = ""
                "state"                = ""
                "zip"                  = ""
                "dateHired"            = "1901-01-01T00:00:00"
                "dateLeft"             = "1901-01-01T00:00:00"
                "phone"                = ""
                "phone2"               = ""
                "comments"             = ""
                "commissionRate"       = 0.00
                "commissionDiscounRate"= $null
                "commissionType"       = ""
                "dateOfBirth"          = "1901-01-01T00:00:00"
                "reviewDate"           = "1901-01-01T00:00:00"
                "password"             = $null
                "resetPassword"        = $null
                "passwordExpiryDate"   = $null
                "active"               = 1
                "allowAdmin"           = 1
                "country"              = ""
                "roleId"               = $roleID
                "tokenId"              = $null
                "employeeType"         = 1
                "Locations"            = $locationIDs
            }

            # Convert the $body hashtable to a JSON-formatted string
            $jsonBody = $body | ConvertTo-Json -Depth 10

            # Make API call
            $response = Invoke-RestMethod -Uri $reqUrl -Method Post -Headers $headers -ContentType 'application/json' -Body $jsonBody

            # Add response to the array
            $responses += $response | Select-Object *
            $responses

            # Display the response content in the console
            Write-Host "`nAPI Response:`n"
            $response | Format-List | Out-Host

            # Export responses to a CSV file after the loop
            $ResponseFileName = "C:\Users\Public\EmployeeDetails-$email-$roleID-$($locationIDs -join '-')-Result.csv"
            $responses | Export-Csv -Path $ResponseFileName -NoTypeInformation

            # Success message
            Write-Host "`nScript executed successfully. Check the CSV file for details.`n"

            # Prompt user to continue or exit
            $userChoice = Read-Host "`nDo you want to enter another set of data? (Y/N)"
            $continue = $userChoice -eq 'Y' -or $userChoice -eq 'y'
        } else {
            Write-Host "`nData not pushed to the API. Returning to input prompt.`n"
            $continue = $true
        }

    } while ($continue)

} catch {
    # Error handling
    Write-Host "`nError: $_`n"
}
