function Import-OutlookMail {
<#
    .SYNOPSIS
    Connects to the user's Outlook account and retrieves all mail from the specified mail folder.

    .DESCRIPTION
    Creates a connection to the logged-in user's microsoft exchange server via instatiation of an
    Outlook.Application Com Object. This Com Object provides an interface known as MAPI that permits 
    the access to, and processing of, all Outlook emails from all Outlook mail folders belonging to the user. 
    
    When invoked with parameters defining which Outlook account and mail folder to retrieve, 
    filters down to the corresponding Outlook account and mail folder returns the entire 
    collection of all mail found in the specified mail folder.
    
    When the Outlook account and the mail folder are not both provided as parameters, a WPF GUI will 
    pop up and allow the user to select which outlook account and mail folder they prefer. Once 
    submitted, those selections will be used, instead of the params, to filter down the collection
    of Outlook accounts and mail folders.

    .PARAMETER OutlookAccountByParam
    [string] - The specific Outlook account to retrieve. If not provided, a WPF GUI is launched
    to retrieve it as input from the user.   

    .PARAMETER MailFolderByParam
    [string] - The specific Mail Folder (i.e. Inbox, Outbox, etc.) from which to retrieve all mail items. 
    If not provided, a WPF GUI is launched to retrieve it as input from the user.
    
    .OUTPUTS
    [__ComObject] - a collection of all emails retrieved from the specified account and mail folder.

    .EXAMPLE
    Import-Outlook

    Note:
    ***An Outlook Account Selection GUI will open and prompt the user for the account and mail folder***

    (ALL EMAILS within the *selected folder* on the *selected account* will be returned after selection)
    
    .EXAMPLE
    Import-Oulook -OutlookAccountByParam user@mail.com -MailFolderByParam Inbox
    
    Note:
    ***No GUI is launched since the required information is provided by the parameters***
    
    (ALL EMAILS within the *selected folder* on the *selected account* are returned)
    
    .EXAMPLE
    Import-Outlook -MailFolderByParam Inbox | sort SentOn -Descending | select -First 1

    Note:
    ***An Outlook Account Selection GUI will open and prompt the user for the account and mail folder***

    (The most recent email within the *selected mail folder* of the *selected account* will be returned)
#>
    
    [CmdletBinding()]
    
    param (
        [string]
        $OutlookAccountByParam,
        
        [string]
        $MailFolderByParam
    )


    begin {

        Write-Verbose "
        Fetching Outlook Accounts from your local Outlook Application's MAPI Namespace
        "

        $OutlookAccounts = (
            New-Object -ComObject outlook.application
        ).GetNamespace('MAPI').Folders
            
        
        Write-Verbose "
        The Outlook Accounts which were retrieved are : $(
            ($OutlookAccounts | Select-Object *).Name -join ', '
        )
        "


        Write-Verbose "
        Defining subroutine to assist with launching the Outlook account selection GUI.
        "
        
        function LaunchOutlookAccountSelectionForm {
            
            param (
                [string]
                $xamlFilePath
            )

            Write-Verbose "
            Now executing LaunchOutlookAccountSelectionForm. The XAML File path is: $xamlFilePath.
            "


            Import-WPFDataFromXAML `
                -xamlFilePath $xamlFilePath

            Write-Verbose "
            After running Import-WPFDataFromXAML, the following variables are now 
            available at the script level: $((Get-Variable |? Name -match 'var_').Name -join ', ')
            "
            

            $OutlookAccounts | ForEach-Object { 
                
                $itemAdded = $var_cbx_Accounts.Items.Add(
                    $_.Name
                )

                Write-Verbose "
                The account $($_.Name) was successfuly added to the account selection list.
                "

            }
            

            $var_cbx_Accounts.Add_SelectionChanged({ HandleSelectionChange })
            $var_cbx_Folders.Add_SelectionChanged({ HandleSelectionChange })
            
            Write-Verbose "
            Added SelectionChanged handlers to $($var_cbx_Accounts.Name) and $($var_cbx_Folders.Name)
            " 
            
            
            $var_btn_Confirm.Add_Click({ HandleClick })
            $form.Add_Closed({ HandleClosed })
            
            Write-Verbose "
            Added Click handler to $($var_btn_Confirm.Name) and Added Closed handler to $($form.Title).
            " 
            
            
            $form.ShowDialog() | Out-Null
            
            Write-Verbose "
            The selection GUI was successfully exited.
            " 

        }
                    

        Write-Verbose "
        Defining subroutine to assist with handling the event when the window is closed.
        "

        function HandleClosed {

            Set-Variable `
            -Name SelectionCanceled `
            -Value $($ConfirmButtonClicked ? $false : $true) `
            -Scope Script

            Write-Verbose " 
            $($SelectionCanceled ? "The account selection was canceled." : "An account was selected.")
            "

        }


        Write-Verbose "
        Defining subroutine to assist with handling the event when a selection changes.
        "

        function HandleSelectionChange{
            
            Set-Variable `
            -Name SelectedOutlookAccount `
            -Value $var_cbx_Accounts.SelectedItem `
            -Scope Script
            
            Set-Variable `
            -Name SelectedMailFolder  `
            -Value $var_cbx_Folders.SelectedItem `
            -Scope Script

            Write-Verbose "
            The selected outlook account is $SelectedOutlookAccount 
            and the selected mail folder is $SelectedMailFolder
            "
        

            $OutlookAccount = $OutlookAccounts | Where-Object Name -match  $SelectedOutlookAccount 

            Write-Verbose "
            The Outlook Account retrieved, ($($OutlookAccount.Name)), 
            matches the selection, ($SelectedOutlookAccount).
            "
            
            
            $OutlookAccount.Folders | ForEach-Object {

                $itemAdded = $var_cbx_Folders.Items.Add($_.Name) 
                
                Write-Verbose "
                The mail folder $($_.Name) was successfuly added to the mail folder selection list.
                "
            }

   
            if( $SelectedOutlookAccount ){  

                $var_lbl_Folder.Visibility = 'Visible'
                $var_cbx_Folders.Visibility = 'Visible'

                Write-Verbose "
                $($var_lbl_Folder.Name) and $($var_cbx_Folders.Name) have now been set to visible.
                "

            }
    

            if( $SelectedMailFolder ){ 

                $var_img_Confirm.Visibility = 'Visible'
                $var_btn_Confirm.Visibility = 'Visible'
  
                Write-Verbose "
                $($var_img_Confirm.Name) and $($var_btn_Confirm.Name) have now been set to visible.
                "

            }

        }


        Write-Verbose "
        Defining subroutine to assist with handling the event when the confirm button is clicked.
        "

        function HandleClick {

            Set-Variable `
                -Name ConfirmButtonClicked `
                -Value $true `
                -Scope Script

            $form.Close()

            Write-Verbose "
            The confirm button was clicked and the selection GUI was closed.
            "
        }

    }
    

    process {

        Write-Verbose "
        Checking to see whether or not if any params have been passed in.
        "

        if( (-not $OutlookAccountByParam) -and (-not $MailFolderByParam) ){

            Write-Verbose "
            The function was invoked without parameters. 
            Now retrieving the Account Selection Form's  XAML file.
            The account selection GUI will now open to query the user 
            for their Outlook account and mail folder selections. 
            "

            LaunchOutlookAccountSelectionForm `
                -xamlFilePath $(
                    Get-ChildItem `
                        -Path $PSScriptRoot `
                        -Filter OutlookAccountSelectionForm_MainWindow.xaml `
                        -Recurse `
                        | 
                        Select-Object -ExpandProperty FullName
                )
        }

        else {

            $SelectedOutlookAccount = $OutlookAccountByParam
            $SelectedMailFolder     = $MailFolderByParam
            
            Write-Verbose "
            The function was invoked with params which already specify the 
            Outlook account and mail folder, so the account selection GUI is
            not needed in this case and has therefore been skipped.
            "

        }


        $OutlookAccount = $OutlookAccounts | Where-Object Name -match $SelectedOutlookAccount       
        
        $Mailbox = $OutlookAccount.Folders | Where-Object Name -match $SelectedMailFolder 
        
        Write-Verbose "
        The Outlook account retrieved after filtering by the account name ($SelectedOutlookAccount) 
        is $($OutlookAccount.Name).
        
        The mail folder retrieved after filtering by the account name ($SelectedOutlookAccount) 
        is $($Mailbox.Name).
        "

    }
    

    end {

        if( -not $SelectionCanceled ){

            
            Write-Verbose "
            Now collecting ALL $SelectedMailFolder messages 
            belonging to the $SelectedOutlookAccount account. 
            "

            $Mailbox.Items
            
            Write-Verbose "
            Collected $(($Mailbox.Items).Length) $(
                ($Mailbox.Items).Length -eq 1 ? "email":  "emails" 
            ).
            "

        }

        else {

            Write-Verbose "
            No messages will be returned because the user has closed 
            the outlook account selection form before providing any input.
            " 
            
        }

    }
}

Export-ModuleMember -Function Import-OutlookMail




function Import-WPFDataFromXAML {
<#
    .SYNOPSIS
    Loads a WPF form using the XMAL file at specified file path and creates script-scoped
    variables which reference all the form's controls by name.   

    .DESCRIPTION
    Adds WPF's PresentationFramework assembly, imports an externally developed WPF Form
    encapsulated within an XAML file, loads that XAML file to convert it into a 
    System.Windows.Window object (which is essentially a WPF Form), then iteratively parses
    the XAML to programatically set variables, each prepended with a 'var_', that capture 
    anything in the XAML found have a 'Name' identifier.

    These variables are set at a script scope. Consequently, they are accessible to the 
    other functions within the module once initialized. 

    .PARAMETER xamlFilePath
    [string ] - The file path to the XAML file that contains the WPF Form. 

    .EXAMPLE
    Import-WPFDataFromXAML $xamlFIlePath

    **NO OUTPUT**

    .OUTPUTS
    Although no immeadiate output is generated as a result of invoking the function,
    There are going to be zero or more variables declared corresponding to the XAML
    schema's inclusion of any named controls. 
    
    These variables will be conventionally named as 'var_ctr_Name', 
    and can be retrieved using the following command:

        Get-Variable | Where-Object -name -match 'var_' 
#>

    [CmdletBinding()]

    param (
        [string]
        $xamlFilePath
    )

    begin{
        # Import WPF
        Add-Type -AssemblyName PresentationFramework
    
        # Extract the Raw Input XAML Content which resides within the file at the given path
        $inputXAML = Get-Content $xamlFilePath -Raw
    
        # Transorm the Input XAML to normalize it for integration with PowerShell 
        $inputXAML = $inputXAML -replace 'mc:Ignorable="d"','' -replace 'x:N','N' -replace '^<Win.*', '<Window'
    }

    process{
        # Convert the Input XAML into a strictly typed XAML object
        [XML] $XAML = $inputXAML
        
        # Parse the XAML object's nodes and store the results in the reader variable
        $reader = New-Object System.Xml.XmlNodeReader $XAML
    
        # Load the reader variable into the XamlReader to generate the form 
        try {
            $form = [Windows.Markup.XamlReader]::Load($reader)
        }
        catch {
            Write-Host $_.Exception
            throw
        }
    }

    end{
        # Iterate through the XAML nodes to create script-scoped variables featuring names,
        # each prepended by 'var_', corresponding to each of the named XML controls  
        $xaml.SelectNodes('//*[@Name]') | ForEach-Object {
            try {
                Set-Variable `
                    -Name "var_$($_.Name)"  `
                    -Value $form.FindName($_.Name) `
                    -Scope Script
            }
            catch {
                Write-Host $_.Exception
                throw
            }
        }

        # Create a script-scoped variable that stores the entire Form 
        Set-Variable -Name "form" -Value $form -Scope Script
    }             
}

Export-ModuleMember -Function Import-WPFDataFromXAML