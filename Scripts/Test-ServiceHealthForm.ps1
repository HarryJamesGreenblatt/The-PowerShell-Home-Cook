function Test-ServiceHealthForm {

    # Import the System.Windows.Forms Namespace
    Add-Type -AssemblyName System.Windows.Forms




    # Instantiate a Form Object
    $FormObject = [System.Windows.Forms.Form]

    # Instantiate a Label Object
    $LabelObject = [System.Windows.Forms.Label]

    # Instantiate a ComboBox Object
    $ComboBoxObject = [System.Windows.Forms.ComboBox]


 

    # Set up the Form:

    #Instantiate a new Form Object
    $ServiceHealthForm = New-Object $FormObject

    #Set the size of the GUI window to 500w x 300h
    $ServiceHealthForm.ClientSize = '500,300'

    #Set the Title of the Form
    $ServiceHealthForm.Text = 'Service Health Form'

    #Set the background color of the form to goldenrod
    $ServiceHealthForm.BackColor = 'whitesmoke'

    #Set the foreground color of the form to white
    $ServiceHealthForm.ForeColor = 'black'



    
    # Add the Dropdown:

    #Instantiate a new Drop Down Object
    $DropDown = New-Object $ComboBoxObject

    #Set the Drop Down's width to 350
    $DropDown.width = 350

    #Set the Drop Down's Placeholder text
    $DropDown.Text = 'Pick a service'

    #Translate the Drop Down to the upper left of form area
    $DropDown.Location =  New-Object System.Drawing.Point(100, 20)

    # Add all the currently running Services into the Drop Down
    Get-Service -ErrorAction SilentlyContinue | 
        Foreach-Object {$DropDown.Items.AddRange(@($_.Name))} 

        
    # Invoke a callback function corresponding to the DropDown's on change event
    # which occurs when the selected index is changed
     $DropDown.Add_SelectedIndexChanged({Set-ServiceDetails})




    # Define a callback function which retreives the value of the currently selected 
    # DropDown item and then sets the Common Name and Status labels using 
    # data pulled from it's Get-Service object
    function Set-ServiceDetails
    {
        $ServiceName           = $DropDown.SelectedItem
        $Details               = Get-Service $Servicename
        $NameValueLabel.Text   = $Details.DisplayName
        $StatusValueLabel.Text = $Details.Status

        $StatusValueLabel.ForeColor = 
            $StatusValueLabel.Text -eq 'Running' ? 'green' : 'red'
    }
        



    # Add the Service Key Text Label:

    #Instantiate a new Label Object
    $ServiceKeyLabel = New-Object $LabelObject

    #Set the Service Key Label Text
    $ServiceKeyLabel.Text = 'Service:'

    #Set the font styling for the Label Text
    $ServiceKeyLabel.Font = 'helvetica,11'

    #Autosize to fit the full message into the label
    $ServiceKeyLabel.AutoSize = $true

    #Translate the Label
    $ServiceKeyLabel.Location = New-Object System.Drawing.Point( 20, 22 )




    # Add the Name Key Text Label:

    #Instantiate a new Label Object
    $NameKeyLabel = New-Object $LabelObject

    #Set the Service Key Label Text
    $NameKeyLabel.Text = 'Common Name:'

    #Set the font styling for the Label Text
    $NameKeyLabel.Font = 'helvetica,11'

    #Autosize to fit the full message into the label
    $NameKeyLabel.AutoSize = $true

    #Translate the Label
    $NameKeyLabel.Location = New-Object System.Drawing.Point( 20, 80 )




    # Add the Name Value Text Label:

    #Instantiate a new Label Object
    $NameValueLabel = New-Object $LabelObject

    #Set the font styling for the Label Text
    $NameValueLabel.Font = 'helvetica,11'

    #Autosize to fit the full message into the label
    $NameValueLabel.AutoSize = $true

    #Translate the Label
    $NameValueLabel.Location = New-Object System.Drawing.Point( 200, 80 )




    # Add the Status Key Text Label:

    #Instantiate a new Label Object
    $StatusKeyLabel = New-Object $LabelObject

    #Set the Service Key Label Text
    $StatusKeyLabel.Text = 'Current Status:'

    #Set the font styling for the Label Text
    $StatusKeyLabel.Font = 'helvetica,11'

    #Autosize to fit the full message into the label
    $StatusKeyLabel.AutoSize = $true

    #Translate the Label
    $StatusKeyLabel.Location = New-Object System.Drawing.Point( 20, 150 )




    # Add the Status Value Text Label:

    #Instantiate a new Label Object
    $StatusValueLabel = New-Object $LabelObject

    #Set the font styling for the Label Text
    $StatusValueLabel.Font = 'helvetica,11'

    #Autosize to fit the full message into the label
    $StatusValueLabel.AutoSize = $true

    #Translate the Label
    $StatusValueLabel.Location = New-Object System.Drawing.Point( 200, 150 )


    

    # Add the Controls to the Form:
    $ServiceHealthForm.Controls.AddRange(
        @(
            $ServiceKeyLabel, 
            $NameKeyLabel, 
            $NameValueLabel, 
            $StatusKeyLabel, 
            $StatusValueLabel, 
            $DropDown
        )
    )
    
    
    
    
    # Display the Form:
    $ServiceHealthForm.ShowDialog()
    
    
    
    
    # Clean up resources after completion of the Form:
    $ServiceHealthForm.Dispose()
    
}

Test-ServiceHealthForm



