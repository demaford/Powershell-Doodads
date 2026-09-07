<#
.SYNOPSIS
    Synchronizes the specified directory's contact list into M365 Outlook personal contact folders.
.DESCRIPTION
    This script is designed to automate the process of synchronizing a structured directory of contact lists (in CSV format) into Microsoft 365 Outlook personal contact folders.
    This script scans the specified directory as the root directory (doesn't)

    The directory structure is scanned recursively, and for each subdirectory, a corresponding contact folder is created in each user's mailbox under a specified root folder.
    For each `contactList.csv` file found in the directory structure, the script imports the contacts into the corresponding contact folder in each user's mailbox.
    The `contactList.csv` file is expected to be in the Outlook contact export format.

    The graph API permissions that are required for this script to function correctly are listed in the notes section.
    While a user can be assigned these permissions, and it would appear that it could apply to all users, only the current user's data is accessible.
    To be able to manage all user's contact folders, the script must authenticate with a service principle or managed identity with he same permissions listed above.


    It supports multiple authentication methods, interactive user and service principles are supported:
    - Interactive User Login (Default)
    - Managed Identity (System or User Assigned)
    - Service Principle (ClientId, TenantId, ClientSecret)
    For details on how to authenticate with each method, please refer to the help examples and the parameter help.
.ROLE
    This script is intended to be run by an administrator or service principle with the necessary permissions to access and modify users' contact folders in Microsoft 365.
.FUNCTIONALITY
    The script performs the following functions:
    1. Authenticates to Microsoft Graph using the specified method.
    2. Retrieves a list of all users in the tenant with a mailbox.
    3. For each user, it checks for the existence of a root contact folder with the specified name.
    4. If the root folder exists, it is deleted to ensure a clean slate for synchronization.
    5. The script then creates a new root contact folder and recursively creates sub-folders based on the directory structure of the specified path.
    6. It imports contacts from CSV files named 'contactList.csv' found in the directory structure and adds them to the corresponding contact folders in each user's mailbox.
.EXAMPLE
    .\Sync-ContactList.ps1 -Path 'C:\ContactLists' -RootFolderName 'Company Contacts'
    This command will synchronize the contact lists found in the "C:\ContactLists" directory into each user's Outlook personal contact folders under a root folder named "Company Contacts".
    This will use the default interactive user login method for authentication.
.EXAMPLE
    .\Sync-ContactList.ps1 -Path 'C:\ContactLists' -RootFolderName 'Company Contacts' -Identity
    This command will synchronize the contact lists found in the "C:\ContactLists" directory into each user's Outlook personal contact folders under a root folder named "Company Contacts".
    This will use the system-assigned managed identity for authentication.
.EXAMPLE
    .\Sync-ContactList.ps1 -Path 'C:\ContactLists' -RootFolderName 'Company Contacts' -Identity -UserAssignedManagedIdentity "your-managed-identity-id"
    This command will synchronize the contact lists found in the "C:\ContactLists" directory into each user's Outlook personal contact folders under a root folder named "Company Contacts".
    This will use the specified user-assigned managed identity for authentication.
.EXAMPLE
    .\Sync-ContactList.ps1 -Path 'C:\ContactLists' -RootFolderName 'Company Contacts' -ClientId "your-client-id" -TenantId "your-tenant-id" -ClientSecret (ConvertTo-SecureString -AsPlainText -Force -String "your-client-secret")
    This command will synchronize the contact lists found in the "C:\ContactLists" directory into each user's Outlook personal contact folders under a root folder named "Company Contacts".
    This will use the service principle authentication method with the provided ClientId, TenantId, and ClientSecret.
.PARAMETER ClientId
    Only when used with service principle authentication.
    The Client ID is the Application ID of the Application Registration that you want to authenticate this script with.
.PARAMETER TenantId
    Only when used with service principle authentication.
    The Tenant ID of the application registration that you want to authenticate this script with.
.PARAMETER ClientSecret
    Only when used with service principle authentication.
    The client secret of the application registration that you want to authenticate this script with.
    This should be passed as a secure string to avoid exposing the secret in plain text.

    Example of how to create a secure string for the client secret:
    ConvertTo-SecureString -AsPlainText -Force -String "clientSecretHere"
.PARAMETER Identity
    Flag that indicates the script should authenticate using a managed identity.
.PARAMETER UserAssignedManagedIdentity
    Only when used with managed identity authentication.
    The ID of the User Assigned Managed Identity that you want to authenticate this script with.
.PARAMETER Path
    The root directory path where the script will begin scanning for contact list CSV files.
    The script will recursively search through this directory and its subdirectories for files named 'contactList.csv'.
    This folder is used as the chroot for execution, and the script will not traverse above this directory.
.PARAMETER RootFolderName
    Display Name of the managed root contact folder that will be created in each user's mailbox.
    All child folders will be created under this root folder, and all contacts will be imported into the appropriate child folders based on the directory structure of the specified path.
.INPUTS
    System.String
    System.Guid
    System.Security.SecureString
    System.IO.DirectoryInfo
.OUTPUTS
    Void
.LINK
    https://github.com/elliot-huffman/Powershell-Doodads
.NOTES
    Requires:
        - The Microsoft.Graph PS Module (version 2.39.0 or later)
        - Graph API Permissions (user or application):
            - Contacts.ReadWrite
                - Allows the user principle to create, read, update, and delete the signed-in user's contacts.
                - Allows the service principle or managed identity to create, read, update, and delete contacts in all users' mailboxes in the organization.
            - User.Read.All
                - Allows the principle to read the full set of profile properties, reports, and managers of other users in your organization.
            - MailboxSettings.Read
                - Allows the user principle to read the signed-in user's mailbox settings.
                - Allows the service principle or managed identity to read the mailbox settings of all users in the organization.
        - PowerShell 7.0 or later
#>

# Define system requirements for powershell pre-launch checks to validate
#Requires -Version 7
#Requires -Module @{ ModuleName="Microsoft.Graph.Authentication"; ModuleVersion="2.39.0" }
#Requires -Module @{ ModuleName="Microsoft.Graph.Users"; ModuleVersion="2.39.0" }
#Requires -Module @{ ModuleName="Microsoft.Graph.PersonalContacts"; ModuleVersion="2.39.0" }

# Enable advanced integrations from PowerShell to act as a native cmdlet
[CmdletBinding(
    SupportsShouldProcess = $true,
    DefaultParameterSetName = 'Default'
)]

# Allow dynamic input for increased usability/flexibility.
param(
    [Parameter(ParameterSetName = 'ServicePrinciple', Mandatory = $true)]
    [System.Guid]$ClientId,
    [Parameter(ParameterSetName = 'ServicePrinciple', Mandatory = $true)]
    [System.Guid]$TenantId,
    [Parameter(ParameterSetName = 'ServicePrinciple', Mandatory = $true)]
    [System.Security.SecureString]$ClientSecret,
    [Parameter(ParameterSetName = 'ManagedIdentity', Mandatory = $true)]
    [Switch]$Identity,
    [Parameter(ParameterSetName = 'ManagedIdentity')]
    [System.Guid]$UserAssignedManagedIdentity,
    [Parameter(ParameterSetName = 'Default')]
    [Parameter(ParameterSetName = 'ManagedIdentity')]
    [Parameter(ParameterSetName = 'ServicePrinciple')]
    [ValidateScript({ Test-Path -Path $_ -PathType 'Container' })]
    [System.IO.DirectoryInfo]$Path = './',
    [Parameter(ParameterSetName = 'Default')]
    [Parameter(ParameterSetName = 'ManagedIdentity')]
    [Parameter(ParameterSetName = 'ServicePrinciple')]
    [System.String]$RootFolderName = 'Organization Managed Contacts'
)

# Setup the components of the script to support the lifecycle of execution
begin {
    # Log in with a service principle instead of the end user's context. This is useful for automation scenarios where user interaction is not desired.
    if ($PSCmdlet.ParameterSetName -eq 'ManagedIdentity') {
        # Disable WAM to support auth types that are not interactive
        Set-MgGraphOption -DisableLoginByWAM $true

        # Check if a user assigned managed identity is provided, and log in with it. If not, use the system-assigned managed identity.
        if ($UserAssignedManagedIdentity) {
            # Log into the MSFT Graph API using a user-assigned managed identity
            Connect-MgGraph -Identity -ClientId $UserAssignedManagedIdentity -NoWelcome
        } else {
            # Log into the MSFT Graph API using the system-assigned managed identity
            Connect-MgGraph -Identity -NoWelcome
        }
    } elseif ($PSCmdlet.ParameterSetName -eq 'ServicePrinciple') {
        # Disable WAM to support auth types that are not interactive
        Set-MgGraphOption -DisableLoginByWAM $true

        # Put the client secret into the correct format for the login command
        [System.Management.Automation.PSCredential]$ClientSecretCredential = New-Object -TypeName 'System.Management.Automation.PSCredential' -ArgumentList $ClientId, $ClientSecret

        # Log into the MSFT Graph API
        Connect-MgGraph -TenantId $TenantId -ClientSecretCredential $ClientSecretCredential -NoWelcome
    } else {
        # Log into the MSFT Graph API
        Connect-MgGraph -Scopes 'Contacts.ReadWrite', 'User.Read.All', 'MailboxSettings.Read' -NoWelcome
    }

    # Mapping used to translate CSV column names to Graph API contact properties.
    # The default mapping is for contacts exported from Outlook, but this can be modified to support other CSV formats as needed.
    #
    # Duplicate entries are allowed for cases where multiple CSV columns map to the same Graph API property.
    # This is useful for scenarios where different CSVs may have different naming conventions for the same data.
    #
    # Where the key is the CSV column name.
    # Where the value is the Graph API contact object's property name.
    [System.Collections.Hashtable]$ContactPropertyMapping = @{
        'Title'                   = 'title'
        'First Name'              = 'givenName'
        'Middle Name'             = 'middleName'
        'Last Name'               = 'surname'
        'Suffix'                  = 'generation'
        'Company'                 = 'companyName'
        'Department'              = 'department'
        'Job Title'               = 'jobTitle'
        'Initials'                = 'initials'
        'Profession'              = 'profession'
        'Notes'                   = 'personalNotes'
        'Business Phone'          = 'businessPhones'
        'Business Phone 2'        = 'businessPhones'
        'Mobile Phone'            = 'mobilePhone'
        'Home Phone'              = 'homePhones'
        'Home Phone 2'            = 'homePhones'
        'Car Phone'               = 'homePhones'
        'E-mail Address'          = 'emailAddresses'
        'E-mail Display Name'     = 'emailAddresses'
        'E-mail Type'             = 'emailAddresses'
        'E-mail 2 Address'        = 'emailAddresses'
        'E-mail 2 Display Name'   = 'emailAddresses'
        'E-mail 2 Type'           = 'emailAddresses'
        'E-mail 3 Address'        = 'emailAddresses'
        'E-mail 3 Display Name'   = 'emailAddresses'
        'E-mail 3 Type'           = 'emailAddresses'
        "Assistant's Phone"       = 'businessPhones'
        'Company Main Phone'      = 'businessPhones'
        'Business Street'         = 'businessAddress'
        'Business Street 2'       = 'businessAddress'
        'Business Street 3'       = 'businessAddress'
        'Business City'           = 'businessAddress'
        'Business State'          = 'businessAddress'
        'Business Postal Code'    = 'businessAddress'
        'Business Country/Region' = 'businessAddress'
        'Home Street'             = 'homeAddress'
        'Home Street 2'           = 'homeAddress'
        'Home Street 3'           = 'homeAddress'
        'Home City'               = 'homeAddress'
        'Home State'              = 'homeAddress'
        'Home Postal Code'        = 'homeAddress'
        'Home Country/Region'     = 'homeAddress'
        'Other Street'            = 'otherAddress'
        'Other Street 2'          = 'otherAddress'
        'Other Street 3'          = 'otherAddress'
        'Other City'              = 'otherAddress'
        'Other State'             = 'otherAddress'
        'Other Postal Code'       = 'otherAddress'
        'Other Country/Region'    = 'otherAddress'
    }
}

# Business logic of the script
process {
    # List of all users in the tenant with a mailbox
    [Microsoft.Graph.PowerShell.Models.MicrosoftGraphUser[]]$AllUsers = Get-MgUser -Property 'Id' -All

    # List of users with their mailbox settings enriched with additional information
    # 
    # Where the Key is the ObjectID of the user
    # Where the value is the mailbox settings of the user
    [System.Collections.Hashtable]$EnrichedUserList = @{}

    # Iterate through each user and get their mailbox settings, storing the results in the hashtable
    foreach ($CurrentUser in $AllUsers) {
        # Retrieve the mailbox settings for the current user
        [Microsoft.Graph.PowerShell.Models.MicrosoftGraphMailboxSettings]$CurrentSettings = Get-MgUserMailboxSetting -UserId $CurrentUser.Id -ErrorAction 'SilentlyContinue'

        # Assign the current settings to the enriched user list if the data is available
        if ($null -ne $CurrentSettings) { $EnrichedUserList[$CurrentUser.Id] = $CurrentSettings }
    }

    # Remove the reference to the all user list to allow garbage collection to free up memory as needed
    $AllUsers = @()

    # List of contact CSVs to be processed into the user's contact folders.
    [System.IO.FileInfo[]]$ContactListTree = Get-ChildItem -Path $Path -Recurse -File -Include 'contactList.csv'

    # We build a unique list of DirectoryInfo objects that represent the tree structure required.
    [System.IO.DirectoryInfo[]]$RequiredDirectories = @()

    # Iterate through each discovered CSV file
    foreach ($File in $ContactListTree) {
        # Start at the directory containing the CSV
        $CurrentDir = $File.Directory
        
        # Walk up the object hierarchy using .Parent until we reach the root $Path
        # We use FullName comparison to ensure we stop exactly at the target root
        while (
            ($null -ne $CurrentDir) -and
            ($CurrentDir.FullName.TrimEnd([System.IO.Path]::DirectorySeparatorChar) -ne $Path.FullName.TrimEnd([System.IO.Path]::DirectorySeparatorChar))
        ) {
            # Do not add duplicate directories to the list.
            if (
                $CurrentDir.FullName.TrimEnd([System.IO.Path]::DirectorySeparatorChar) -notin
                @($RequiredDirectories | ForEach-Object { $_.FullName.TrimEnd([System.IO.Path]::DirectorySeparatorChar) })
            ) { $RequiredDirectories += $CurrentDir }
            
            # Recursively move up to the parent directory for the next iteration
            $CurrentDir = $CurrentDir.Parent
        }
    }

    # Sort the directories by their depth (number of segments in FullName).
    # This ensures we process "Root/Sub" before "Root/Sub/Leaf", allowing top-down creation.
    [System.IO.DirectoryInfo[]]$SortedDirectories = $RequiredDirectories | Sort-Object { $_.FullName.Split([System.IO.Path]::DirectorySeparatorChar).Count }

    # Clear the memory reference used by the list so that garbage collection can free up memory as needed.
    $RequiredDirectories = @()

    # Sync each user's contact list with the corresponding CSVs found in the specified path
    foreach ($UserId in $EnrichedUserList.Keys) {
        # Query for the specified root folder
        [Microsoft.Graph.PowerShell.Models.MicrosoftGraphContactFolder]$OldRootContactFolder = Get-MgUserContactFolder -UserId $UserId -Filter "displayName eq '$RootFolderName'"

        # Check if WhatIf/Confirm is specified, implement risk mitigation
        if ($PSCmdlet.ShouldProcess('Outlook Contact Folders', "Delete Root Managed Folder: $($OldRootContactFolder.DisplayName)")) {
            # If the root folder exists, remove it to ensure a clean slate for synchronization. This is done to avoid conflicts and ensure that the contact list is accurately represented based on the CSVs provided.
            if ($OldRootContactFolder -is [Microsoft.Graph.PowerShell.Models.MicrosoftGraphContactFolder]) { Remove-MgUserContactFolderPermanent -UserId $UserId -ContactFolderId $OldRootContactFolder.Id }
        }
        
        # Check if WhatIf/Confirm is specified, implement risk mitigation
        if ($PSCmdlet.ShouldProcess('Outlook Contact Folders', "Create New Managed Root Folder: $RootFolderName")) {
            # Newly created root folder that will act as the base for the new contact list structure.
            [Microsoft.Graph.PowerShell.Models.MicrosoftGraphContactFolder]$NewRootContactFolder = New-MgUserContactFolder -UserId $UserId -DisplayName $RootFolderName
        }
        
        # Track the current position in the Graph API contact folder tree for this specific user
        [System.String]$CurrentGraphParentId = $NewRootContactFolder.Id

        # 2. Folder Creation Logic (Top-Down Traversal)
        foreach ($Directory in $SortedDirectories) {
            # Check if a folder with this name already exists within the current parent scope
            [Microsoft.Graph.PowerShell.Models.MicrosoftGraphContactFolder]$ExistingChild = Get-MgUserContactFolderChildFolder -UserId $UserId -ContactFolderId $CurrentGraphParentId | Where-Object -FilterScript { $_.DisplayName -eq $Directory.Name }

            # If the folder does not exist, create it and update the pointer to this new folder ID. If it does exist, simply move the pointer to this existing folder to continue the traversal downwards.
            if ($null -eq $ExistingChild) {
                # Check if WhatIf/Confirm is specified, implement risk mitigation
                if ($PSCmdlet.ShouldProcess('Outlook Contact Folders', "Create Child Folder: $($Directory.Name)")) {
                    # Folder does not exist in Graph: Create it and update the pointer to this new folder ID
                    [Microsoft.Graph.PowerShell.Models.MicrosoftGraphContactFolder]$NewSubFolder = New-MgUserContactFolderChildFolder -UserId $UserId -DisplayName $Directory.Name -ContactFolderId $CurrentGraphParentId
                    
                    # Update the current parent ID to the newly created folder's ID for the next iteration
                    $CurrentGraphParentId = $NewSubFolder.Id
                }

            } else {
                # Folder exists: Move the pointer to this existing folder to continue the traversal downwards
                $CurrentGraphParentId = $ExistingChild.Id
            }

            # Check if a contact list is present in the current directory
            [System.IO.FileInfo]$ContactFileList = Get-ChildItem -Path $Directory.FullName -File -Filter 'contactList.csv'

            # Check if a contact list was found, and if so, import it to the current folder
            if ($ContactFileList) {
                # Import the contact list CSV into a custom object for processing
                [PSCustomObject]$ContactList = Import-Csv -Path $ContactFileList.FullName

                # Iterate through each contact and create it in the user's contact folder using the Graph API.
                foreach ($CsvContact in $ContactList) {
                    # Object that will be splatted into the graph API contact creation cmdlet.
                    [System.Collections.Hashtable]$ContactCreationBody = @{}

                    # We iterate through every property present on the CSV row to ensure 100% coverage.
                    foreach ($CsvColumnName in $CsvContact.PSObject.Properties.Name) {
                        $CsvValue = $CsvContact.$CsvColumnName
                        
                        # Rule: If the incoming param is null or "" then it should be excluded from the splat.
                        if ([System.String]::IsNullOrWhiteSpace($CsvValue)) { continue }

                        # Determine target Graph property name using your provided mapping; fallback to Column Name if not mapped.
                        [System.String]$GraphProp = $ContactPropertyMapping.ContainsKey($CsvColumnName) ? $ContactPropertyMapping[$CsvColumnName] : ''

                        # Skip un-mapped columns to avoid type errors on submission
                        if ($GraphProp -eq '') { continue }

                        # --- EMAIL ADDRESSES ---
                        elseif ($GraphProp -eq 'emailAddresses') {
                            # Skip current column if the column name does not contain "Address"
                            if ($CsvColumnName -notmatch 'Address') { continue }

                            # Find the corresponding Type column for this specific email entry (e.g., "E-mail Address Type")
                            [System.String]$DisplayNameColumnName = $CsvColumnName -replace 'Address', 'Display Name'

                            # Display name for the current email address
                            [System.String]$DisplayName = $CsvContact.$DisplayNameColumnName

                            # Ensure the collection exists in the splat
                            if (-not $ContactCreationBody.ContainsKey('emailAddresses')) { $ContactCreationBody['emailAddresses'] = @() }

                            # Construct the emailAddress object
                            [System.Collections.Hashtable]$NewEmail = @{
                                'address' = $CsvValue
                                'name'    = $DisplayName
                            }
                            
                            # Add to collection without overwriting previous emails (E-mail 1, E.g., Email 2)
                            $ContactCreationBody.$GraphProp += $NewEmail
                        }

                        # --- PHONE COLLECTIONS ---
                        elseif ($GraphProp.EndsWith('Phones')) {
                            # Ensure the collection exists as an array [string[]]
                            if (-not $ContactCreationBody.ContainsKey($GraphProp)) { $ContactCreationBody[$GraphProp] = @() }

                            # Add to the existing array (prevents overwriting Business Phone 1 with Business Phone 2)
                            $ContactCreationBody[$GraphProp] += $CsvValue
                        }

                        # --- PHYSICAL ADDRESSES ---
                        elseif ($GraphProp -in @('businessAddress', 'homeAddress', 'otherAddress')) {
                            # Initialize the address object if not present
                            if (-not $ContactCreationBody.ContainsKey($GraphProp)) {
                                # Template to inject into the new property to be added to the creation body
                                [System.Collections.Hashtable]$SubBodyObject = @{
                                    'street'          = ''
                                    'city'            = $null
                                    'state'           = $null
                                    'postalCode'      = $null
                                    'countryOrRegion' = $null
                                }

                                # Add the new property to the creation body
                                $ContactCreationBody[$GraphProp] = $SubBodyObject
                            }

                            # Map CSV sub-segments to Graph PhysicalAddress properties
                            if ($CsvColumnName -match 'Street') { $ContactCreationBody.$GraphProp['street'] = $ContactCreationBody.$GraphProp['street'] -eq '' ? $CsvValue : "$($ContactCreationBody.$GraphProp['street']) $CsvValue" }
                            elseif ($CsvColumnName.EndsWith('City')) { $ContactCreationBody.$GraphProp['city'] = $CsvValue }
                            elseif ($CsvColumnName.EndsWith('State')) { $ContactCreationBody.$GraphProp['state'] = $CsvValue }
                            elseif ($CsvColumnName.EndsWith('Postal Code')) { $ContactCreationBody.$GraphProp['postalCode'] = $CsvValue }
                            elseif ($CsvColumnName.EndsWith('Country/Region')) { $ContactCreationBody.$GraphProp['countryOrRegion'] = $CsvValue }
                        }

                        # --- STANDARD STRING PROPERTIES ---
                        else { $ContactCreationBody[$GraphProp] = $CsvValue }
                    }

                    # Check if WhatIf/Confirm is specified, implement risk mitigation
                    if ($PSCmdlet.ShouldProcess("Outlook Contact Folder: $($Directory.Name)", "Create Contact: $($ContactCreationBody['givenName']) $($ContactCreationBody['surname'])")) {
                        # Create the new contact object with the provided properties
                        New-MgUserContactFolderContact -UserId $UserId -ContactFolderId $CurrentGraphParentId -BodyParameter $ContactCreationBody | Out-Null
                    }
                }
            }
        }
    }
}

# Cleanup operations
end {
    # Log out of the Graph API
    Disconnect-MgGraph | Out-Null
}