function Import-WPFDataFromXML {
    param (
        [string]
        $xamlFilePath
    )
    begin{
        # Import WPF
        Add-Type -AssemblyName PresentationFramework
    
        # Extract the Raw Input XML Content which resides within the file at the given path
        $inputXAML = Get-Content $xamlFilePath -Raw
    
        # Transorm the Input XML to normalize it for integration with PowerShell 
        $inputXAML = $inputXAML -replace 'mc:Ignorable="d"','' -replace 'x:N','N' -replace '^<Win.*', '<Window'
    }
    process{
        # Convert the Input XAML into a strictly typed XML object
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



function Test-HelloWorldXMLForm {

    Import-WPFDataFromXML '../WPF Forms/HelloWorldForm_MainWindow.xaml'

    function ToggleHello {
        if($var_lbl_HelloWorld.Text -eq '')
        {
            $var_lbl_HelloWorld.Text = 'Hello World'
        }  
        else
        {
            $var_lbl_HelloWorld.Text = ''
        }
    }


    $var_btn_Toggle.Add_Click({ToggleHello}) 


    $form.ShowDialog()
    
}


Test-HelloWorldXMLForm