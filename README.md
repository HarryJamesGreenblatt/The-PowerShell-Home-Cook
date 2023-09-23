# The-PowerShell-Home-Cook
You know, if PowerShell was a musical instrument, I think it would be a Conductor's Baton  :heavy_minus_sign::ok_hand:

## Modules
### [`Search-Internet`](./Modules/Search-Internet/Search-Internet.psm1)
Searches the web for the given Query String on an optionally specified Search Engine and Browser.   
   
<br>   
   
### [`Write-Figlet`](./Modules/Write-Figlet/Write-Figlet.psm1)
Converts a provided string input to generative ASCII Typesetting produced by the python pyfiglet module.
   
<br>   
   
### [`Import-Outlook`](./Modules/Import-OutlookMail/Import-Outlook.psm1)
A collection of function which facilitate the management of Outlook messages using Outlook's  Message Application Programming Interface (MAPI).
   
- #### [`Import-WPFDataFromXAML`](./Modules/Import-Outlook/Import-Outlook.psm1)
    Loads a WPF form using the XMAL file at specified file path and creates script-scoped variables which reference all the form's controls by name. 

- #### [`Receive-OutlookMail`](./Modules/Import-Outlook/Import-Outlook.psm1)
    Connects to the user's Outlook account and retrieves the specified mail folder as a powershell Com Object.   

- #### [`Limit-OutlookMail`](./Modules/Import-Outlook/Import-Outlook.psm1)
    Limits the quantity of ComObjects collected from the user's Outlook Mail Folder using a DASL Filter Query string.   

<br>
<br>

## Scripts
### [`Test-HelloWorldForm.ps1`](./Scripts/Test-HelloWorldForm.ps1)
Uses Windows Forms to generate a simple button GUI using 100% vanilla powershell
   
<br>

### [`Test-HelloWorldXMLForm.ps1`](./Scripts/Test-HelloWorldXMLForm.ps1)
Loads an externally developed WPF Form saved in XAML format into a simple button GUI whose control logic is handled by PoweerShell.
   
<br>

### [`Test-ServiceHealthForm.ps1`](./Scripts/Test-ServiceHealthForm.ps1)
Uses Windows Forms to generate a PowerShell GUI that implements a Combo Box to list the Health of Services currently running on the user's machine.