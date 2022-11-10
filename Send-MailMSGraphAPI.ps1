function Send-MailMSGraphAPI {
    param (
        [Parameter(Mandatory = $true)]
        [String]
        $TenantId,
        [Parameter(Mandatory = $true)]
        [String]
        $ClientId,
        [Parameter(Mandatory = $true)]
        [String]
        $ClientSecret,
        [Parameter(Mandatory = $true)]
        [String]
        $From,
        [Parameter(Mandatory = $true)]
        [String]
        $To,
        [Parameter(Mandatory = $true)]
        [String]
        $Subject,
        [Parameter(Mandatory = $true)]
        [String]
        $Body,
        [Parameter(Mandatory = $false)]
        [Boolean]
        $SaveToSentItems
    )
    Write-Host $SaveToSentItems
    $TokenBody = @{
        Grant_Type    = "client_credentials"
        Scope         = "https://graph.microsoft.com/.default"
        Client_Id     = $ClientId
        Client_Secret = $ClientSecret
    }
    $TokenResponse = Invoke-RestMethod -Uri "https://login.microsoftonline.com/$TenantId/oauth2/v2.0/token" -Method POST -Body $TokenBody
    $Headers = @{
        "Authorization" = "Bearer $($TokenResponse.access_token)"
        "Content-Type"  = "application/json"
    }
    $SendURL = "https://graph.microsoft.com/v1.0/users/$From/sendMail"
    $SendBody = @{
        "message"         = @{
            "subject"      = $Subject
            "body"         = @{
                "contentType" = "HTML"
                "content"     = $Body
            }
            "toRecipients" = @(
                @{
                    "emailAddress" = @{
                        "address" = $To
                    }
                }
            )
        }
        "saveToSentItems" = $SaveToSentItems
    }
    $SendJsonBody = $SendBody | ConvertTo-Json -Depth 4
    Invoke-RestMethod -Method POST -Uri $SendURL -Headers $Headers -Body $SendJsonBody
}
