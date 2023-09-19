function Test-ServiceHealthForm {

    # Import the System.Windows.Forms Namespace
    Add-Type -AssemblyName System.Windows.Forms



    # Instantiate a Form Object
    $FormObject = New-Object System.Windows.Forms.Form

    # Instantiate a Label Object
    $LabelObject = New-Object System.Windows.Forms.Label

    # Instantiate a Button Object
    $ButtonObject = New-Object System.Windows.Forms.Button

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




    # Add the Hello World Text to the Label:

    #Instantiate a new Label Object
    $Label = $LabelObject

    #Set the Label text to store the Hello Word message
    $Label.Text = 'Service Health'

    #Autosize to fit the full message into the label
    $Label.AutoSize = $true

    #Translate the Label
    $Label.Location = New-Object System.Drawing.Point( ((500/2) - 100), ((300/2) - 50) )

    #Set the font styling
    $Label.Font = 'helvetica,24,style=Bold'




    #instantiate a new Button Object
    $Button = $ButtonObject

    #Set the Button's text label
    $Button.Text = 'Push this button'

    $Button.Font= 'Helvetica,14'

    #Autosize the button to esure proper formatting
    $Button.Autosize = $true

    #Translate the button underneath the message text
    $Button.Location =  New-Object System.Drawing.Point(165, 155)

    


    # Define a Callback Function 
    function ToggleHello 
    {
        if($Label.Text -eq '')
        {
            $Label.Text = 'Service Health'
        }
        
        else
        {
            $Label.Text = ''
        }
    }

    # Assign the Callback Function to the Button's Click event
    $Button.Add_Click({ToggleHello})



    
    # Add the Controls to the Form:
    $ServiceHealthForm.Controls.AddRange(@($Label, $Button))
    
    
    
    
    # Display the Form:
    $ServiceHealthForm.ShowDialog()
    
    
    
    
    # Clean up resources after completion of the Form:
    $ServiceHealthForm.Dispose()
    

}

Test-ServiceHealthForm



