function Receive-OutlookMailbox {
<#
    .SYNOPSIS
    Connects to the user's Outlook account and retrieves the specified mail folder as a powershell Com Object.

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
    [__ComObject] - a reference to the specified account and mail folder.

    .EXAMPLE
    Receive-OutlookMailbox

    Note:
    ***An Outlook Account Selection GUI will open and prompt the user for the account and mail folder***

    (ALL EMAILS within the *selected folder* on the *selected account* will be returned after selection)
    
    .EXAMPLE
    Import-Oulook -OutlookAccountByParam user@mail.com -MailFolderByParam Inbox
        
        Note:
        ***No GUI is launched since the required information is provided by the parameters***
        
        (ALL EMAILS within the *selected folder* on the *selected account* are returned)
    
    .EXAMPLE
    Receive-OutlookMailbox -MailFolderByParam Inbox | sort SentOn -Descending | select -First 1

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

            $Mailbox
            
        }

        else {

            Write-Verbose "
            No messages will be returned because the user has closed 
            the outlook account selection form before providing any input.
            " 
            
        }

    }
}

Export-ModuleMember -Function Receive-OutlookMailbox




function Limit-OutlookMailbox {
<#
    .SYNOPSIS
    Limits the quantity of ComObjects collected from the user's Outlook Mail Folder 
    using a DASL Filter Query string.

    .DESCRIPTION
    Given a ComObject referencing an Outlook mail folder via the MAPI namespace,
    the initial state is a reference to the full contents of the mail folder,
    which often are quite large (thousands of messages). 
    
    As such, executing sort or filter operations via the pipeline,
    
        ex. $OutlookMailbox | Sort-Object SentOn | Where-Object SenderName -match google 

    is extremely slow and inefficient due to the data's existance as a reference to an external
    data source, which, as a ComObject, is essentially asking the exchange server to remotely
    sort on the ENTIRE contents of the mailbox, then search the ENTIRE contents of the mailbox again
    for the search string, and then return the result to powershell at the end, for thousands of 
    messages, each of which track dozens of properties.
    
    One approach which can be used to mitigate this limitation is the Restrict() method inherent to
    the Outlook ComObject's 'Items' property, which is inherent to each Outlook Folder.

    The Restrict() Method takes a DASL Filter Query String which employs uniquely specific syntax.
    Properly filtering the data source using these DASL filter queries results in markedly improved
    performance for both filter and search operations. Below is the equivalent DASL to reproduce the 
    example given above:
    
        ex. $OutlookMailbox.Items.Restrict( SQL@="urn:schemas:httpmail:sender" LIKE '%google%' )

    The function will automatically convert a less verbose string into this syntax, such that running
    the function with the following arguments will acheive the same result as the example above:
    
        ex. $OutlookMailbox | Limit-OutlookMailbox "sender LIKE '%google%'"

    .PARAMETER FilterQuery
    [string] - A simplified , albeit DASL-compliant, filter query string which only requires
    the   ur:schemas:httpmail endpoint,   the operator   and the   serach term.
    
    definitively,
                " sender                    LIKE                   '%google%' "
        urn:schemas:httpmail endpoint,   the operator   and the    serach term.

    .PARAMETER Mailbox
    [__ComObject] - A reference connecting to the entire contents of an Outlook mail folder.
    May be optionally passed as pipeline input. 

    .EXAMPLE
    Receive-OutlookMailbox $OutlookAcoount $MailFoler | Limit-OutlookMailbox "sender LIKE '%google%'"

        ( Returns only the messages in the $MailFolder sent by google. )
    
    .OUTPUTS
    [__ComObject] - a filtered view of the specified mail folder limited to the messages
    described by the DASL query
#>

    [CmdletBinding()]

    param (

        [Parameter(Mandatory=$true)]
        [string] 
        $FilterQuery,

        [Parameter(ValueFromPipeline, Mandatory=$true)]
        [__ComObject] 
        $Mailbox

    )
    

    begin {
        
        Write-Verbose "
        The unmodified Filter Query provided was:`n`n`t$FilterQuery
        "

        if( ( $FilterQuery | Select-String '\sAND|OR\s' ).Matches.Success ){

            $DASL =  $FilterQuery `
                -replace '(\bfrom\b|\bto\b|\bsubject\b|\bdatereceived\b|\btextdescription\b|\bhasattachment\b|\battachmentfilename\b|\bsender\b|\bcc\b|\bbcc\b|\breply-to\b|\bpriority\b|\bread\b|\breferences\b|\bthread-topic\b|\bthread-index\b)', 'urn:schemas:httpmail:$1' `
                -replace '(urn:schemas:httpmail:[a-z]+)' , '"$1"' `
                -replace '(.+)'                          , '@SQL=$1' `

        }
        
        else{

            $DASL = $FilterQuery `
                -replace '(^.+\s)' , 'urn:schemas:httpmail:$1' `
                -replace '(.+)', '@SQL=$1' `
                -replace '(urn:schemas:httpmail:[a-z]+)' , '"$1"' `

        }
        
        Write-Verbose "
        After converting the Filter Query to conform to DASL syntax, we have:
        `n`t$DASL
        "

    }
    

    process {
        
        Write-Verbose "
        Now limiting the $($Mailbox.Name) mail folder only to 
        messages that satisfy the DASL Query.
        "
        
        $FilteredMail   =   $Mailbox.Items.Restrict( $DASL )

        $MailSize       = ( $FilteredMail | Select-Object * ).Length 

    }
    

    end {

        $FilteredMail
        
        Write-Verbose "
        The  $($Mailbox.Name)  mail folder 
        was successfully limited to:   $MailSize messages.
        "

    }

}

Export-ModuleMember -Function Limit-OutlookMailbox





function Send-OutlookMail {
<#
    .SYNOPSIS
    Sends an Outlook message, inculding any attachments, by way of 
    either a plaintext or HTML Body format, to the given recipient address.

    .DESCRIPTION
    Creates a COM Object reference to the Outlook Application, which provides
    a template to submit a New Email. Once the COM Object reference is established, 
    resulting in the creation of a MailItem Object Reference, each of the 
    given parameters are stored into a hash object, where they are then
    accessed iteratively by their Keys (which are named after the same properties present
    in the MailItem Object). This operation sets the New Mail Item's properties to reflect 
    the prameters values provided.

    The message can be formatted either as plaintext, which is acheivable by way 
    of providing a string input into the Body parameter, or as HTML, which is 
    acheivable by way of providing an HTML string into the HTML Body parameter. 
    
    Once the New Mail Item's properties get set, an optional Disply switch paramter 
    is evaluated. If given, the Display switch will then open the Outlook application's
    New Email window, already prefilled with the values (To, Subject, Body, etc) 
    which were provided as parameters. Note however that, in this scenario, the user 
    must manually send the message.
    
    Otherwise, should the Display switch parameter not be given, the Outlook message is
    sent programatically, without manual intervention.
    
    .PARAMETER To
    [string] - The recipients email address. 

    .PARAMETER Subject
    [string] - the subject of the Outlook message.
    
    .PARAMETER Body
    [string] - A plaintext string that will serve as the Outlook message body.

    .PARAMETER HTMLBody
    [string] - An HTML string which, if provided, will be rendered as the 
    Outlook message body. 

    .PARAMETER Display
    [switch] - A boolean flag that initiates launching the Outlook Application
    in order to manually send the message. 

    .EXAMPLE
    Send-OutlookMail `
        -To         'Them@Email.com '`
        -Subject    'Message Subject Goes Here' `
        -Body       'Plaintext message body.'

        ( programatically sends a plaintext body message to given recipient  )
        
    .EXAMPLE
    Send-OutlookMail `
        -To         'Them@Email.com '`
        -Subject    'Message Subject Goes Here' `
        -HTMLBody   '<h1> HTML message body. </h1>'
        -Display
        
        ( opens the Outlook App's New Mail window, and allows an HTML body message 
          to the given recipient to be sent manually  )

    .EXAMPLE
    Send-OutlookMail `
        -To            'Them@Email.com '`
        -Subject       'Message Subject Goes Here' `
        -Body          'Plaintext message body.'
        -Attachments   @('./File1.pdf', '../../File2.xlsx')

        ( programatically sends a plaintext body message to given recipient
          which includes all the files given as attachments )

#>

    [CmdletBinding()]
    
    param (
        
        [Parameter(Mandatory=$true)]
        [string]
        $To,

        [string]
        $Subject,
        
        [string]
        $Body,
        
        [string] 
        $HTMLBody,

        [string[]]
        $Attachments,

        [switch] 
        $Display

    )

    
    begin {

        
        $Outlook = New-Object -ComObject Outlook.Application
        
        
        Write-Verbose "
        Now creating a New Mail Object in Outlook and storing the
        given parameters in a hash object. 
        "
        
        $NewMail = $Outlook.CreateItem(0)

        $params = @{
            To          = $To
            Subject     = $Subject
            HTMLBody    = $HTMLBody
            Body        = $Body
            Attachments = $Attachments
        }

        Write-Verbose "
        The New Mail Item has been created and is of type $($NewMail.GetType()).`n
        The params hash object has been created and has stored the following keys:
        $(($params.Keys | Sort-Object ) -join ', ').  
        "

    }
    

    process {

        Write-Verbose "
        Now iterating through the keys of the given parameters to
        assign them as values within the New Mail Object 
        "

        $params.Keys | Sort-Object | ForEach-Object { 

            If( $_ -eq 'Attachments'){
                
                $params[ 'Attachments' ] | ForEach-Object {
                    $NewMail.Attachments.Add($_) | Out-Null
                }

                Write-Verbose "
                $_ was added as an Attachment to the New Mail Item.
                "

            }
            
            else {

                $NewMail.$_ = $params[ $_ ] 

                Write-Verbose "
                The $_ Property in the New Mail Object has now been set to:  
                $($params[ $_ ]) 
                "

            }
        
        }
        
    }
    

    end {

        Write-Verbose "
        Determining whether or not to Display the New Mail Window.
        "
        
        if( $Display ){

            Write-Verbose "
            The Display switch was given. Now launching the New Mail Window in Outlook.
            "
            $NewMail.GetInspector.Activate()
        
        }

        else {

            $NewMail.Send()

            Write-Verbose "
                $(
                    $NewMail.Sent `
                    ? `
                    "The Email to $($NewEmail.To) was successfully sent." `
                    : `
                    "The Email was not sent. Something went wrong."
                )
            "
            
        }

        Exit-OutlookSession
    
    }

}

Export-ModuleMember -Function Send-OutlookMail





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





function Exit-OutlookSession {
<#
    .SYNOPSIS
    Stops all Outlook running Outlook processes as to not leave persistent connections
    to the Exchange Server.

    .DESCRIPTION
    Invokes Get-Process to collect any running instances of the Outlook Application,
    then pipes those instances into a command to stop the process.
#>

    Get-Process -Name Outlook | Stop-Process

}

Export-ModuleMember -Function Exit-OutlookSession
