function Test-ServiceHealthForm {

    # Import the System.Windows.Forms Namespace
    Add-Type -AssemblyName System.Windows.Forms




    # Instantiate a Form Object
    $FormObject = New-Object System.Windows.Forms.Form

    # Instantiate a Label Object
    $LabelObject = New-Object System.Windows.Forms.Label

    # Instantiate a ComboBox Object
    $ComboBoxObject = New-Object System.Windows.Forms.ComboBox


 

    # Set up the Form:

    #Instantiate a new Form Object
    $ServiceHealthForm = $FormObject

    #Set the size of the GUI window to 500w x 300h
    $ServiceHealthForm.ClientSize = '500,300'

    #Set the Title of the Form
    $ServiceHealthForm.Text = 'Service Health Form'

    #Set the background color of the form to goldenrod
    $ServiceHealthForm.BackColor = 'darkblue'

    #Set the foreground color of the form to white
    $ServiceHealthForm.ForeColor = 'white'




    # Add the Text Label:

    #Instantiate a new Label Object
    $Label = $LabelObject

    #Set the Label text to store the Hello Word message
    $Label.Text = 'Service:'

    #Set the font styling for the Label Text
    $Label.Font = 'helvetica,11'


    #Autosize to fit the full message into the label
    $Label.AutoSize = $true

    #Translate the Label
    $Label.Location = New-Object System.Drawing.Point( 20, 22 )




    # Add the Dropdown:

    #Instantiate a new Drop Down Object
    $DropDown = $ComboBoxObject

    #Set the Drop Down's width to 350
    $DropDown.width = 350

    #Set the Drop Down's Placeholder text
    $DropDown.Text = 'Pick a service'

    #Translate the Drop Down to the upper left of form area
    $DropDown.Location =  New-Object System.Drawing.Point(100, 20)

    # Add all the currently running Services into the Drop Down
    Get-Service | Foreach-Object {$DropDown.Items.AddRange(@($_.Name))} 



    
    # Add the Controls to the Form:
    $ServiceHealthForm.Controls.AddRange(@($Label, $DropDown))
    
    
    
    
    # Display the Form:
    $ServiceHealthForm.ShowDialog()
    
    
    
    
    # Clean up resources after completion of the Form:
    $ServiceHealthForm.Dispose()
    
}

Test-ServiceHealthForm



