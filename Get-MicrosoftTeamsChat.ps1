#Requires -Version 7.0
<#

    .SYNOPSIS
        Exports Microsoft Chat History

    .DESCRIPTION
        This script reads the Microsoft Graph API and exports of chat history into HTML files in a location you specify.

    .PARAMETER ExportFolder
        Export location of where the HTML files will be saved. For example, "D:\ExportedHTML\"

    .PARAMETER clientId
        The client id of the Azure AD App Registration. Not required if using a Microsoft Graph session.

    .PARAMETER tenantId
        The domain name of the UPNs for users in your tenant. E.g. contoso.com. Not required if using a Microsoft Graph session.
    
    .PARAMETER domain
        The heritage tenant, Readify or Kloud. Not required if using a Microsoft Graph session.


    .EXAMPLE
        .\Get-MicrosoftTeamChat.ps1 -ExportFolder "D:\ExportedHTML" -clientId "ClientIDforAzureADAppRegistration" -tenantId "TenantIdoftheAADOrg" -domain "contoso.com"

    .NOTES
        Author:  Trent Steenholdt
        Pre-requisites: An app registration with delegated User.Read, Chat.Read and User.ReadBasic.All permissions is needed in the Azure AD tenant you're connecting to.

#>

[cmdletbinding()]
Param(
    [Parameter(Mandatory = $true, HelpMessage = "Export location of where the HTML files will be saved.")] [string] $ExportFolder,
    [Parameter(Mandatory = $false, HelpMessage = "The client id of the Azure AD App Registration")] [string] $clientId,
    [Parameter(Mandatory = $false, HelpMessage = "The tenant id of the Azure AD environment the user logs into")] [string] $tenantId,
    [Parameter(Mandatory = $false, HelpMessage = "The domain name of the UPNs for users in your tenant. E.g. contoso.com")] [string] $domain
)

#################################
##   Import Modules  ##
#################################

Set-Location $PSScriptRoot

Import-Module ($PSScriptRoot + "/functions/TelstraPurpleFunctions") -Force

Get-TPASCII

#region Build HTML
####################################
##   HTML  ##
####################################

$HTML = Get-Content -Raw ./files/chat.html

$HTMLMessagesBlock_them = @"
<div class="message-container">
    <div class="message">
        <div style="display:flex; margin-top:10px">
            <div style="flex:none; overflow:hidden; border-radius:50%; height:42px; width:42px; margin:10px">
                <img height="42" src="###IMAGE###" style="vertical-align:top; width:42px; height:42px;" width="42">
            </div>
            <div class="them" style="flex:1; overflow:hidden;">
                <div style="font-size:1.2rem; white-space:nowrap; text-overflow:ellipsis; overflow:hidden;">
                    <span style="font-weight:700;">###NAME###</span><span style="margin-left:1rem;">###DATE###</span>
                </div>
                <div>
                    ###CONVERSATION###
                </div>
                ###ATTACHMENT###
            </div>
        </div>
    </div>
</div>
"@


$HTMLMessagesBlock_me = @"
<div class="message-container">
    <div class="message">
        <div style="display:flex; margin-top:10px">
            <div class="me" style="flex:1; overflow:hidden;">
                <div style="font-size:1.2rem; white-space:nowrap; text-overflow:ellipsis; overflow:hidden;">
                    <span style="font-weight:700;">###NAME###</span><span style="margin-left:1rem;">###DATE###</span>
                </div>
                <div>
                    ###CONVERSATION###
                </div>
                ###ATTACHMENT###
            </div>
            <div style="flex:none; overflow:hidden; border-radius:50%; height:42px; width:42px; margin:10px">
                <img height="42" src="###IMAGE###" style="vertical-align:top; width:42px; height:42px;" width="42">
            </div>
        </div>
    </div>
</div>
"@

$HTMLAttachmentBlock = @"
<div class="attachment">
<a href="###ATTACHMENTURL###" target="_blank">###ATTACHMENTNAME###</a>
</div>
"@
#endregion Build HTML

# Script
#region Connect
Write-Output "`r`nStarting script..."
if ([string]::IsNullOrEmpty($clientId)) {
    Write-Output "`r`nUsing Microsoft Graph session..."
    While ([string]::IsNullOrEmpty((Get-MgContext).Account)) {
        Write-Warning "`r`nNot signed in to MgGraph."
        Write-Output "`r`nLaunching MgGraph sign in..."
        if (Get-Module -Name Microsoft.Graph.Authentication) {
            Connect-MgGraph -Scopes "Chat.Read", "User.Read", "User.ReadBasic.All" -nowelcome
        }
        else {
            Write-Output "`r`nMicrosoft.Graph.Authentication module not found. Please install the Microsoft Graph PowerShell SDK and try again. See https://learn.microsoft.com/en-us/powershell/microsoftgraph/installation?view=graph-powershell-1.0 for directions."
            Exit
        }
    }
    $authtype = "MSGraph"
}
else {
    Write-Host -ForegroundColor White "`r`nSign in with the Device Code to the app registration:"
    $tokenOutput = Connect-DeviceCodeAPI $clientId $tenantId $null
    $token = $tokenOutput.access_token
    $refresh_token = $tokenOutput.refresh_token
    $accessToken = ConvertTo-SecureString $token -AsPlainText -Force
    $authtype = "AppReg"
}
#endregion Connect

#region Prep Folders
$ImagesFolder = Join-Path -Path $ExportFolder -ChildPath 'images'
if (-not(Test-Path -Path $ImagesFolder)) { New-Item -ItemType Directory -Path $ImagesFolder | Out-Null }
$ExportFolder = (Resolve-Path -Path $ExportFolder).ToString()
#endregion Prep Folders

#region Get Current User
if ($authtype -eq "MSGraph") {
    $me = Invoke-MgGraphRequest -Method Get "https://graph.microsoft.com/v1.0/me" 
}
else {
    $me = Invoke-RestMethod -Method Get -Uri "https://graph.microsoft.com/v1.0/me" -Authentication OAuth -Token $accessToken
}
#endregion Get Current User

#region Get Chats
# Get the id, topic, chatType, createdDateTime, and lastUpdatedDateTime of the first set of chats
$allChats = @();
if ($authtype -eq "MSGraph") {
    $firstChat = Invoke-MgGraphRequest -Method Get "https://graph.microsoft.com/v1.0/me/chats?`$Select=id,topic,chatType,createdDateTime,lastUpdatedDateTime" 
}
else {
    $firstChat = Invoke-RestMethod -Method Get -Uri "https://graph.microsoft.com/v1.0/me/chats?`$Select=id,topic,chatType,createdDateTime,lastUpdatedDateTime" -Authentication OAuth -Token $accessToken
}
$allChats += $firstChat
$allChatsCount = $firstChat.'@odata.count' 

Write-Output ("`r`nGetting all chats, please wait... This may take some time.`r`n")

# Get the id, topic, chatType, createdDateTime, and lastUpdatedDateTime of the rest of the chats
if ($null -ne $firstChat.'@odata.nextLink') {
    $chatNextLink = $firstChat.'@odata.nextLink'
    do {
        if ($authtype -eq "MSGraph") {
            $chatsToAdd = Invoke-MgGraphRequest -Method Get $chatNextLink
        }
        else {
            $chatsToAdd = Invoke-RestMethod -Method Get -Uri $chatNextLink -Authentication OAuth -Token $accessToken
        }
        $allChats += $chatsToAdd
        $chatNextLink = $chatsToAdd.'@odata.nextLink'
        $allChatsCount = $allChatsCount + $chatsToAdd.'@odata.count'
    } until ($null -eq $chatsToAdd.'@odata.nextLink' )
}
#endregion Get Chats

# Sort the chats by the date they were created
$chats = $allChats.value | Sort-Object createdDateTime -Descending
Write-Output ("`r`n" + $chats.count + " possible chat threads found.`r`n")

$threadCount = 0
$StartTime = Get-Date


foreach ($thread in $chats) {
    
    # 50 is the maximum allowed with the beta api
    $conversationUri = "https://graph.microsoft.com/v1.0/me/chats/" + $thread.id + "/messages?top=50"

    $elapsedTime = (Get-Date) - $StartTime
    
    Write-Verbose ("Script running for " + $elapsedTime.TotalSeconds + " seconds.")
    
    # Refresh token every 30 minutes if using app registration
    #region Token Refresh
    if ($authtype -eq  "AppReg") {
        if ($elapsedTime.TotalMinutes -gt 30) {
            Write-Host -ForegroundColor Cyan "Reauthenticating with refresh token..."
            $tokenOutput = Connect-DeviceCodeAPI $clientId $tenantId $refresh_token
            $token = $tokenOutput.access_token
            $refresh_token = $tokenOutput.refresh_token
            $accessToken = ConvertTo-SecureString $token -AsPlainText -Force
            $StartTime = $(Get-Date)
            Start-Sleep 5
        }
    }
    #endregion Token Refresh

    # If the thread has a topic, use that as the name of the chat. Otherwise, use the names of the members.
    #region Set Chat Name
    $name = Get-Random;

    if ($null -ne $thread.topic) {
        $name = $thread.topic
        
    }
    else {
        $membersUri = "https://graph.microsoft.com/v1.0/me/chats/" + $thread.id + "/members"
        if ($authtype -eq "MSGraph") {
            $members = Invoke-MgGraphRequest -Method Get $membersUri
        }
        else {
            $members = Invoke-RestMethod -Method Get -Uri $membersUri -Authentication OAuth -Token $accessToken
        }
        $members = $members.value.displayName | Where-Object { $_ -notlike "*@purple.telstra.com" }
        $name = ($members | Where-Object { $_ -notmatch $me.displayName } | Select-Object -Unique) -join ", "
        
    }
    #endregion Set Chat Name

    # Ok ladies now let's get conversations    
    $allConversations = @();
    #region Get Messages
    try {
        if ($authtype -eq "MSGraph") {
            $firstConversation = Invoke-MgGraphRequest -Method Get $conversationUri
        }
        else {
            $firstConversation = Invoke-RestMethod -Method Get -Uri $conversationUri -Authentication OAuth -Token $accessToken
        }
        $allConversations += $firstConversation
        $allConversationsCount = $firstConversation.'@odata.count' 
    }
    catch {
        Write-Output ($name + " :: Could not download historical messages.")
        Write-Host -ForegroundColor Yellow "Skipping...`r`n"
    }

    if ($null -ne $firstConversation.'@odata.nextLink') {
        $conversationNextLink = $firstConversation.'@odata.nextLink'
        do {
            if ($authtype -eq "MSGraph") {
                $conversationToAdd = Invoke-MgGraphRequest -Method Get $conversationNextLink
            }
            else {
                $conversationToAdd = Invoke-RestMethod -Method Get -Uri $conversationNextLink -Authentication OAuth -Token $accessToken
            }
            $allConversations += $conversationToAdd
            $conversationNextLink = $conversationToAdd.'@odata.nextLink'

            $allConversationsCount = $allConversationsCount + $conversationToAdd.'@odata.count'
        } until ($null -eq $conversationToAdd.'@odata.nextLink')
    }
    #endregion Get Messages

    $conversation = $allConversations.value | Sort-Object createdDateTime 
    $threadCount++
    $messagesHTML = $null
      
    if (($conversation.count -gt 0) -and (-not([string]::isNullorEmpty($name)))) {

        Write-Host -ForegroundColor White ($name + " :: " + $allConversationsCount + " messages.")
        Write-Verbose $conversationUri 

        foreach ($message in $conversation) {
            # Commenting out the below line because it assumes a pattern for display name that is not valid for my use case.
            #$userPhotoUPN = ($message.from.user.displayName -replace " ", ".") + "@" + $domain

            # Getting the user's UPN via the id in the message works no matter what the display name pattern is.
            if ($authtype -eq "MSGraph") {
                try {
                    $userPhotoUPN = (Invoke-MgGraphRequest -Method GET "https://graph.microsoft.com/v1.0/users/$($message.from.user.id)?`$Select=userPrincipalName" -erroraction SilentlyContinue)["userprincipalname"]
                }
                catch {
                    Write-Output "Could not get userPrincipalName for userid $($message.from.user.id)"
                }
                
            }
            else {
                $userPhotoUPN = Invoke-RestMethod -Method Get -Uri "https://graph.microsoft.com/v1.0/users/$($message.from.user.id)?`$Select=userPrincipalName" -Authentication OAuth -Token $accessToken
            }
           
            $profilefile = Join-Path -Path $ImagesFolder -ChildPath "$userPhotoUPN.jpg"
            if (-not(Test-Path $profilefile)) {
                $profilePhotoUri = "https://graph.microsoft.com/v1.0/users/" + $userPhotoUPN + "/photos/96x96/`$value"
                $pictureURL = ("data:image/jpeg;base64,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")
                try {
                    if ($authtype -eq "MSGraph") {
                        Invoke-MgGraphRequest -Method Get -Uri $profilePhotoUri -OutputFilePath $profilefile 
                    }
                    else {
                        Invoke-WebRequest -Uri $profilePhotoUri -Authentication OAuth -Token $accessToken -OutFile $profilefile 
                    }
                    $pictureURL = Get-EncodedImage $profilefile
                }
                catch {
                    Write-Verbose "Could not get picture for user"  
                }
            }
            else {
                $pictureURL = Get-EncodedImage $profilefile
            }

            $messageBody = $message.body.content
            if ($messageBody -match "<img.+?src=[\`"']https:\/\/graph.microsoft.com(.+?)[\`"'].*?>") {
                $imagecount = 0

                foreach ($imgMatch in $Matches) {
                    $imagecount++
                    $threadidIO = $thread.id.Split([IO.Path]::GetInvalidFileNameChars()) -join '_'
                    $imagefile = Join-Path -Path $ImagesFolder -ChildPath "$threadidIO$imagecount.jpg"
                    $imageUri = "https://graph.microsoft.com" + $imgMatch[1]

                    Write-Output "Downloading embedded image in message..."
                    Write-Verbose $imageUri

                    $retries = 0
                    $limit = 5
                    $completed = $false
                    while (-not $completed) {
                        try {
                            if ($authtype -eq "MSGraph") {
                                $response = Invoke-MgGraphRequest -Method Get -Uri $imageUri -OutputFilePath $imagefile
                            }
                            else {
                                $response = Invoke-WebRequest -Uri  $imageUri -Authentication OAuth -Token $accessToken
                                Set-Content -Path $imagefile -AsByteStream -Value $response.Content
                                if ($response.StatusCode -ne 200) {
                                    throw "Expecting reponse code 200, was: $($response.StatusCode)"
                                }
                            }
                                                        
                            $imageencoded = Get-EncodedImage $imagefile
                            $messageBody = $messageBody.Replace($imgMatch[0], ("<a href=`"" + $imageencoded + "`" download>" + $imgMatch[0] + "</a>"))
                            $messageBody = $messageBody.Replace($imageUri, $imageencoded)
                            
                            $completed = $true
                        }
                        catch {
                            if ($retries -ge $limit) {
                                Write-Warning "Request to $imageUri failed the maximum number of $limit times."
                                $completed = $true
                            }
                            else {
                                Write-Warning "Request to $imageUri failed. Retrying in 5 seconds."
                                Write-Warning "Image file: $imagefile"
                                Start-Sleep 5
                                $retries++
                            }
                        }
                    }
                }
            }

            $time = [System.TimeZoneInfo]::ConvertTimeBySystemTimeZoneId((Get-Date ($message.createdDateTime)), (Get-TimeZone).Id)
            $time = Get-Date $time -Format "dd MMMM yyyy, hh:mm tt"

            if ($message.from.user.displayName -eq $me.displayName) {
                $HTMLMessagesBlock = $HTMLMessagesBlock_me
            } 
            else { 
                $HTMLMessagesBlock = $HTMLMessagesBlock_them
            }
            # Get attachments and add message plus attachment to HTML Messages Block
            if ($null -ne $message.attachment) {
                $attachmentHTML = $HTMLAttachmentBlock `
                    -Replace "###ATTACHMENTURL###", $message.attachment.name`
                    -Replace "###ATTACHMENTNAME###", $message.attachment.contentURL`
                    -Replace "###IMAGE###", $pictureURL

                $messagesHTML += $HTMLMessagesBlock `
                    -Replace "###NAME###", $message.from.user.displayName`
                    -Replace "###CONVERSATION###", $messageBody`
                    -Replace "###DATE###", $time`
                    -Replace "###ATTACHMENT###", $attachmentHTML`
                    -Replace "###IMAGE###", $pictureURL
                    
            }
            # No attachments. Add message to HTML Messages Block
            else {
                $messagesHTML += $HTMLMessagesBlock `
                    -Replace "###NAME###", $message.from.user.displayName`
                    -Replace "###CONVERSATION###", $messageBody`
                    -Replace "###DATE###", $time`
                    -Replace "###ATTACHMENT###", $null`
                    -Replace "###IMAGE###", $pictureURL
            }
        }
        # Build and output the HTML file
        $HTMLfile = $HTML `
            -Replace "###MESSAGES###", $messagesHTML`
            -Replace "###CHATNAME###", $name`

        $name = $name.Split([IO.Path]::GetInvalidFileNameChars()) -join '_' 

        if ($name.length -gt 64) {
            $name = $name.Substring(0, 64);
        }

        $file = Join-Path -Path $ExportFolder -ChildPath "$name.html"
        if (Test-Path $file) { $file = ($file -Replace ".html", ( "(" + $threadCount + ")" + ".html")) }
        Write-Host -ForegroundColor Green "Exporting $file... `r`n"
        $HTMLfile | Out-File -FilePath $file
        
        # Enter the thread info to the CSV chat list
        # Export to CSV barfs if the value for topic is null. So we need to check for and fix that, maybe by replacing it with the name generated above, although we'd need to be sure that was populated too.
        if ($null -eq $thread.topic) {
            $thread.topic = $name
        }
        $thread | Export-Csv "$ExportFolder\ChatThreads.csv" -NoTypeInformation -Append
        Write-output $thread
    }
    else {
        Write-Output "$name :: No messages found."
        Write-Verbose "Thread ID: $($thread.id)"
        Write-Host -ForegroundColor Yellow "Skipping...`r`n"
    }
}
Remove-Item -Path $ImagesFolder -Recurse
Write-Host -ForegroundColor Cyan "`r`nScript completed... Bye!"
