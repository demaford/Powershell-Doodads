<#
.SYNOPSIS
    Find and replace text on all word documents in the specified directory.
.DESCRIPTION
    Will do a case insensitive find and replace in the specified word documents.
    The end user has the option to enable recursive behavior if desired.
.EXAMPLE
    PS C:\> Show-WordFindReplaceUI.ps1
    Runs the script in user interface mode.
    The user can then select their required options using a Windows.Forms UI.
.PARAMETER Path
    The path parameter is the directory that contains the word docs or other folders that needs to be recurse into for the find and replace operation.
    If you want recurse behavior, please specify the recurse parameter.
.PARAMETER Recurse
    When specified, the find and replace operation will search in sub folders of the specified path.
    When this parameter is omitted, the find and replace operation will only operate on the specified path and will not go into child directories.
.PARAMETER CLIMode
    When this parameter is invoked, this forces the UI to not be rendered and only to use the command line.
    This is useful if you want an automated operation that has no human interaction.
.INPUTS
    System.String
.OUTPUTS
    Output (if any)
.NOTES
    Requires .net framework desktop as it uses windows forms to render a UI.
    Exit Codes:
        1 - MS Word has not been initialized properly, check to ensure it has been installed.
        2 - MS Word process closed while script was running and the script could not recover from this.
            Closing MS Word while a COM Object is loaded will clear the COM Object even though the com object is separate from the GUI.
#>

#Requires -Version 5.1
#Requires -PSEdition Desktop

[CmdletBinding(SupportsShouldProcess = $true)]

param(
    [Parameter(
        Position = 0,
        ValueFromPipeline = $true,
        ValueFromPipelineByPropertyName = $true
    )
    ]
    [ValidateScript( {
            Test-Path -Path $_ -PathType "Container"
        })]
    [ValidateNotNullOrEmpty()]
    [System.String]$Path,

    [Parameter(
        Position = 1,
        ValueFromPipeline = $true,
        ValueFromPipelineByPropertyName = $true
    )
    ]
    [ValidateNotNullOrEmpty()]
    [System.String]$Find,

    [Parameter(
        Position = 2,
        ValueFromPipeline = $true,
        ValueFromPipelineByPropertyName = $true
    )
    ]
    [ValidateNotNullOrEmpty()]
    [System.String]$Replace,

    [Switch]$Recurse,
    [Switch]$CLIMode
)

begin {
    # Allow PowerShell to access the Windows Forms name space
    Add-Type -AssemblyName "System.Windows.Forms"

    # Enable the Windows theming engine to theme the windows form that is rendered by PowerShell
    [System.Windows.Forms.Application]::EnableVisualStyles()

    function New-MSWord {
        <#
        .SYNOPSIS
            Returns a Word COM object instance
        .DESCRIPTION
            Initializes Word and returns a COM instance of it after error checking.
        .EXAMPLE
            PS C:\> New-MSWord
            Return an Word COM object.
        .OUTPUTS
            Microsoft.Office.Interop.Word.ApplicationClass
        .NOTES
            Requires Word to be installed on the computer.

            Exit Codes:
            1 - Word has not been initialized properly, check to ensure it has been installed.
        #>

        Write-Verbose -Message "Instantiating Word object"

        # Initialize Word
        $WordObject = New-Object -ComObject "Word.application"

        # Write debug info to the console
        Write-Debug -Message $WordObject

        # Check to see if the object has been created properly
        if ($WordObject -IsNot [Microsoft.Office.Interop.Word.ApplicationClass]) {
            # Write an error message to stderr (this is non-terminating)
            Write-Error "Word has not been initialized properly. Check to make sure it is installed."

            # Return $False for a failed initialization
            $PSCmdlet.WriteObject($false)

            # Exit Script execution unsuccessfully
            Exit 1
        }
        else {
            # If the object was created, return it
            Return $WordObject
        }
    }

    function Show-DirectoryBrowserUI {
        <#
        .SYNOPSIS
            Displays the directory tree selector for the end user to select a folder.
        .DESCRIPTION
            Launches the directory tree selector and returns the path of the selected folder as a string.
            If the user cancels, returns boolean $False.
        .EXAMPLE
            PS C:\> Show-DirectoryBrowserUI
            Launches the directory tree selector and returns the path of the selected folder as a string.
            If the user cancels, returns boolean $False.
        .INPUTS
            None
        .OUTPUTS
            System.String
            System.Boolean
        .NOTES
            Will return false if the user cancels.
        #>

        # Create an directory file dialog box and set the accepted file types.
        $OpenDialog = New-Object "System.Windows.Forms.FolderBrowserDialog"
        # Automatically select the currently selected path
        if ($Path) {
            $OpenDialog.SelectedPath = $Path
        }
        else {
            $OpenDialog.SelectedPath = ""
        }

        # Show the dialog box and capture the results.
        # If the user selected a file return the file path. Otherwise return false if nothing is selected.
        if ($OpenDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
            Return $OpenDialog.SelectedPath
        }
        else {
            Return $false
        }
    }

    function Show-MainUI {
        <#
        .SYNOPSIS
            Renders the main UI
        .DESCRIPTION
            Builds the main user interface and starts the render process to display the main user interface to the end user.
            The user interface can call the find and replace functions to execute the bulk find and replace operations.
        .EXAMPLE
            PS C:\> Show-MainUI
            Renders the main user interface for end users to interact with.
        .OUTPUTS
            System.Object[]
        .NOTES
            Requires the .Net framework (not core) as it needs to render Windows.Forms applications
        #>

        #Requires -Version 5.1

        # Create the blank form with require
        $Form = New-Object "System.Windows.Forms.Form"
        $Form.ClientSize = New-Object System.Drawing.Point(400, 400)
        $Form.text = "Word Bulk Find and Replace"
        $Form.TopMost = $true

        $FindLabel = New-Object "System.Windows.Forms.Label"
        $FindLabel.text = "Find:"
        $FindLabel.AutoSize = $true
        $FindLabel.width = 25
        $FindLabel.height = 10
        $FindLabel.location = New-Object System.Drawing.Point(17, 18)
        $FindLabel.Font = New-Object System.Drawing.Font('Segoe UI', 12)

        $FindTextBox = New-Object "System.Windows.Forms.TextBox"
        $FindTextBox.multiline = $false
        $FindTextBox.width = 180
        $FindTextBox.height = 20
        $FindTextBox.location = New-Object System.Drawing.Point(17, 45)
        $FindTextBox.Font = New-Object System.Drawing.Font('Segoe UI', 12)
        if ($Find) { $FindTextBox.Text = $Find }

        $ReplaceLabel = New-Object "System.Windows.Forms.Label"
        $ReplaceLabel.text = "Replace With:"
        $ReplaceLabel.AutoSize = $true
        $ReplaceLabel.width = 25
        $ReplaceLabel.height = 10
        $ReplaceLabel.location = New-Object System.Drawing.Point(17, 94)
        $ReplaceLabel.Font = New-Object System.Drawing.Font('Microsoft Sans Serif', 12)

        $ReplaceTextBox = New-Object "System.Windows.Forms.TextBox"
        $ReplaceTextBox.multiline = $false
        $ReplaceTextBox.width = 180
        $ReplaceTextBox.height = 20
        $ReplaceTextBox.location = New-Object System.Drawing.Point(17, 118)
        $ReplaceTextBox.Font = New-Object System.Drawing.Font('Microsoft Sans Serif', 12)
        if ($Replace) { $ReplaceTextBox.Text = $Replace }

        $WorkingDirectoryLabel = New-Object "System.Windows.Forms.Label"
        $WorkingDirectoryLabel.text = "Working Directory:"
        $WorkingDirectoryLabel.AutoSize = $true
        $WorkingDirectoryLabel.width = 25
        $WorkingDirectoryLabel.height = 10
        $WorkingDirectoryLabel.location = New-Object System.Drawing.Point(17, 167)
        $WorkingDirectoryLabel.Font = New-Object System.Drawing.Font('Segoe UI', 12)

        $SelectedDirLabel = New-Object "System.Windows.Forms.Label"
        $SelectedDirLabel.text = "No working directory selected..."
        $SelectedDirLabel.AutoSize = $true
        $SelectedDirLabel.width = 25
        $SelectedDirLabel.height = 10
        $SelectedDirLabel.location = New-Object System.Drawing.Point(17, 191)
        $SelectedDirLabel.Font = New-Object System.Drawing.Font('Segoe UI', 12)
        if ($Path) { $SelectedDirLabel.Text = $Path }

        $BrowseButton = New-Object "System.Windows.Forms.Button"
        $BrowseButton.text = "Browse for Directory"
        $BrowseButton.width = 178
        $BrowseButton.height = 30
        $BrowseButton.location = New-Object System.Drawing.Point(17, 248)
        $BrowseButton.Font = New-Object System.Drawing.Font('Microsoft Sans Serif', 12)

        $RecurseCheckBox = New-Object "System.Windows.Forms.CheckBox"
        $RecurseCheckBox.text = "Recursive"
        $RecurseCheckBox.AutoSize = $false
        $RecurseCheckBox.width = 95
        $RecurseCheckBox.height = 20
        $RecurseCheckBox.location = New-Object System.Drawing.Point(18, 299)
        $RecurseCheckBox.Font = New-Object System.Drawing.Font('Segoe UI', 12)
        if ($Recurse) { $RecurseCheckBox.Checked = $Recurse }

        $FindAndReplaceButton = New-Object "System.Windows.Forms.Button"
        $FindAndReplaceButton.text = "Find And Replace"
        $FindAndReplaceButton.width = 148
        $FindAndReplaceButton.height = 30
        $FindAndReplaceButton.enabled = $true
        $FindAndReplaceButton.location = New-Object System.Drawing.Point(125, 343)
        $FindAndReplaceButton.Font = New-Object System.Drawing.Font('Microsoft Sans Serif', 12)

        # Place the objects onto an array to be able to access them later.
        $UIItems = @(
            $FindLabel,
            $FindTextBox,
            $ReplaceLabel,
            $ReplaceTextBox,
            $WorkingDirectoryLabel,
            $SelectedDirLabel,
            $BrowseButton,
            $RecurseCheckBox,
            $FindAndReplaceButton
        )

        # Add controls to the form to be rendered at render start
        $Form.controls.AddRange($UIItems)

        # Start the form render loop
        [void]$Form.ShowDialog()

        # Return the UI objects so that they can be manipulated at other scopes
        # Return $UIItems
    }

    # Capture the common parameter overrides to inherit the values to all cmdlets
    switch (0) {
        { -not $PSBoundParameters.ContainsKey('Debug') } { $DebugPreference = $PSCmdlet.SessionState.PSVariable.GetValue('DebugPreference') }
        { -not $PSBoundParameters.ContainsKey('Verbose') } { $VerbosePreference = $PSCmdlet.SessionState.PSVariable.GetValue('VerbosePreference') }
        { -not $PSBoundParameters.ContainsKey('Confirm') } { $ConfirmPreference = $PSCmdlet.SessionState.PSVariable.GetValue('ConfirmPreference') }
        { -not $PSBoundParameters.ContainsKey('WhatIf') } { $WhatIfPreference = $PSCmdlet.SessionState.PSVariable.GetValue('WhatIfPreference') }
    }

    # Write debug info to the console
    Write-Debug -Message $VerbosePreference
    Write-Debug -Message $ConfirmPreference
    Write-Debug -Message $WhatIfPreference

    # Initialize the file edited tracker
    [System.Collections.Hashtable[]]$EditedFiles = $null

    # Instantiate MS Word
    $MSWord = New-MSWord
}

process {
    # Check to see if MS Word was terminated while the script was running.
    # If the ComObject is empty, the script will try to re-init the object. If the re-init fails, the script exits unsuccessfully.
    if ($null -eq $MSWord.Application) {

        # Write verbose info to the console
        Write-Verbose -Message "MS Word is not currently initialized, re-initializing MS Word"

        # Re-init the MS Word object
        $MSWord = New-MSWord

        # Check if the MS Word Object is still null.
        if ($null -eq $MSWord.Application) {
            # Write debug info to the console
            Write-Debug -Message $MSWord

            # Write a message to stderr (non-terminating)
            Write-Error -Message "The MS Word application was closed while the script was running!"

            # Return $false for failed operation
            $PSCmdlet.WriteObject($false)

            # Exit script unsuccessfully with exit code
            exit 2
        }
    }
}

end {
    # Stop the MS Word Process
    $MSWord.Quit()

    # Clean up the objects that were created
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($MSWord) | Out-Null
    $MSWord = $Null

    # Check if the script is dot sourced, if it is then do not execute the stuff inside.
    if (($MyInvocation.Line -NotMatch "^\.\s") -and (-not $CLIMode)) {
        # Show the main UI for user interaction
        Show-MainUI
    }
}