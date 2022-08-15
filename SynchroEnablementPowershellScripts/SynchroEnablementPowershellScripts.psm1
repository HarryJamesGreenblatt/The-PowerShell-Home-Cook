function Get-HealthAndStatus {
<#
.SYNOPSIS
        Initiates a remote session between endpoints and packages a collection of health and status metrics 
        and log data belonging to both into a Custom Object.

.DESCRIPTION
        The Client will have it's System Information, derived from the  Get-ComputerInfo Cmdlet, packaged into a
        Hash Table called ClientHealthAndStatus, along with a List summarizing all of its current Secrity Logs. 
        The ClientHealthAndStatus will then be passed as an input to a PowerShell Remoting session concuted between
        the Client and the Server over SSH. 
        
        During the remoting session, the Server will go through a similar process,
        producing it's own Hash table called ClientHealthAndStatus, which will be exported back to the Client following 
        the termination of the remoting session.
        
        Finally, a Custom PSObject is created to store both Hash Tables, effectively enabling to 
        capability of accessing each endpoint's data by it's role and the type of report,  
                
                i.e.  $HealthAndStatus.Server.SystemInfo   or    $HealthAndStatus.Client.LogList 

.PARAMETER UserName
        The Client's User Account Name belonging to the User who may authenticate to the Server.

.PARAMETER HostName
        The "Computer Name" assigned to the Server.

.PARAMETER PathToTransferDirectory
        The Path to a File Transfer Directory located on the Server. 

.NOTES
        May be used concurrently with Write-ToTransferDirectory.
        (See Examples)

.EXAMPLE
        Get-HealthAndStatus `
                UserName = My_Name `
                HostName = Server_Name `
                PathToTransferDirectory = Path\to\directory

.EXAMPLE
        $params = @{
                UserName = My_Name;
                HostName = Server_Name;
                PathToTransferDirectory= Path\to\directory
        }

        Get-HealthAndStatus @params -Verbose

.EXAMPLE
        $params = @{
                UserName = My_Name;
                HostName = Server_Name;
                PathToTransferDirectory= Path\to\directory
        }

        Get-HealthAndStatus @params | Write-ToTransferDirectory @params
#>
        [CmdletBinding()]
        
        param (

                [Parameter(Mandatory=$true)]
                [string]
                $UserName,

                [Parameter(Mandatory=$true)]
                [string]
                $HostName,

                [Parameter(Mandatory=$true)]
                [string]
                $PathToTransferDirectory

        )
        
        begin {

                Write-Verbose "The path to the Server's Transfer Directory is:  $PathToTransferDirectory."
                Write-Verbose "Initiating Session over SSH to  $UserName@$HostName."
                
                $Session = New-PSSession -UserName $UserName -HostName $HostName


                Write-Verbose ("Packaging Client's System Information " +
                "and Security Log Summary Data into a Health and Status Hash Table.")
                
                $ClientHealthAndStatus = @{

                        SystemInfo = Get-ComputerInfo;

                        LogList = Get-WinEvent -ListLog Security | 
                                                        Format-List -Property *  
                                                                
                }
                
                Write-Verbose "The Client's Health and Status is now stored in a:`n  $ClientHealthAndStatus."
                        
        }
        

        process {
                
                Write-Verbose ("Invoking a Command to the Server to package it's System Information" + 
                " and Security Log Summary Data into a seperate Hash Table.")
                
                Invoke-Command `
                        -Session $Session `
                        -OutVariable ServerHealthAndStatus `
                        -ScriptBlock {
                        
                                @{
                                        SystemInfo = Get-ComputerInfo;
                                        
                                        LogList = Get-WinEvent -ListLog Security |
                                                                        Format-List -Property *                        
                                }

                        }

                Exit-PSSession

        }
        
        end {

                Write-Verbose "Returning from the remoting session, the Server's Health and Status is now held in a:`n  $ServerHealthAndStatus."
                Write-Verbose "Packaging both the Client and Server Health and Status Hashes into a Custom PSObject."
                
                [PSCustomObject]@{
                        Client = $ClientHealthAndStatus;
                        Server = $ServerHealthAndStatus
                }
                
                Write-Verbose "The Health and Status Custom PSObject is now available."
        }
}
Export-ModuleMember -Function Get-HealthAndStatus                     





function Write-ToTransferDirectory {
<#
.SYNOPSIS
        Initiates a remote session between endpoints, unpackages a Custom PSObject into a specified directory 
        designated as a staging area to facilitate data transfer, and modifies the unpacked data to remove undesirable 
        non-printing characters.

.DESCRIPTION
        The Client will initiate a remoting session with the Server and check for the existence of a specified Transfer Directory.
        If it doesn't exist, one will be created.

        Following that, the contents of the $HealthAndStatus input, whether passed in as Pipeline Input, or as 
        a Parameter Value, are unpacakged and written as files respective to the System Information and Log Lists included  
        with the Client and Server Health and Status Hash Tables that are returned from Get-HealthAndStatus.
        
        However, given that PowerShell 7 unfortantely includes ANSI Color Codes in its File Representation of it's Objects,
        Additional measures are taken to Remove these ANSI Color Codes using a REGEX Subsitution via  -replace  '',''.

.PARAMETER HealthAndStatus,
        A Custom PSObject containing Hash Tables corresponding to the Client's and Server's Health and Status Data, respectively.

.PARAMETER UserName
        The Client's User Account Name belonging to the User who may authenticate to the Server.

.PARAMETER HostName
        The "Computer Name" assigned to the Server.

.PARAMETER PathToTransferDirectory
        The Path to a File Transfer Directory located on the Server. 

.PARAMETER ANSIColorCodes
        The Regular Expression matching the ANSI Color Code junk characters included by default in PowerShell 7 output. 

.NOTES
        Get-HealthAndStatus is used concurrently with this function.
        (See Examples)

.EXAMPLE
        Write-ToTransferDirectory `
                HealthAndStatus = Get-HealthAndStatus @params `
                UserName = My_Name `
                HostName = Server_Name `
                PathToTransferDirectory = Path\to\directory

.EXAMPLE
        $params = @{
                HealthAndStatus = Get-HealthAndStatus @params
                UserName = My_Name;
                HostName = Server_Name;
                PathToTransferDirectory= Path\to\directory
        }

        Write-ToTransferDirectory @params -Verbose

.EXAMPLE
        $params = @{
                UserName = My_Name;
                HostName = Server_Name;
                PathToTransferDirectory= Path\to\directory
        }

        Get-HealthAndStatus @params | Write-ToTransferDirectory @params
#>
        [CmdletBinding()]
        
        param (
                [Parameter(ValueFromPipeline=$true)]
                [System.Object[]]
                $HealthAndStatus,
        
                [Parameter(Mandatory=$true)]
                [string]
                $UserName,

                [Parameter(Mandatory=$true)]
                [string]
                $HostName,

                [Parameter(Mandatory=$true)]
                [string]
                $PathToTransferDirectory,
                
                [string]
                $ANSIColorCodes = '\x1B\[([0-9]{1,3}((;[0-9]{1,3})*)?)?[m|K]'           
        
        )
        
        begin {

                Write-Verbose "The path to the Server's Transfer Directory is:  $PathToTransferDirectory."
                Write-Verbose "Initiating Session over SSH to  $UserName@$HostName."
                
                $Session = New-PSSession -UserName $UserName -HostName $HostName


                Write-Verbose "Setting up an array of File Names to facilitate File Write operations."
                
                $Deliverables = @(
                        "Client Info",
                        "Client Logs",
                        "Server Info",
                        "Server Logs"
                )   
        }
        
        process {

                Write-Verbose ("Invoking a Command to the Server to unpackage it's Health and Status Object " +
                "into Files stored in the specified Transfer Directory.")
                
                Invoke-Command `
                        -Session $Session `
                        -ArgumentList $HealthAndStatus,$Deliverables,$PathToTransferDirectory,$ANSIColorCodes `
                        -ScriptBlock {

                                Param(
                                        [PSCustomObject]$HealthAndStatus, 
                                        [System.Object[]]$Deliverables,
                                        $PathToTransferDirectory,
                                        $ANSIColorCodes
                                )

                                
                                Write-Verbose "Checking If a Transfer Directory exists."
                                
                                If( -not (Test-Path $PathToTransferDirectory) ){
                                        New-Item -ItemType Directory $PathToTransferDirectory
                                }


                                Write-Verbose "Unpackaging the Health and Status Object into individual Transfer Files."
                                
                                $HealthAndStatus.Client.SystemInfo | Out-File "$PathToTransferDirectory\Client Info" -Encoding ASCII
                                $HealthAndStatus.Client.LogList    | Out-File "$PathToTransferDirectory\Client Logs" -Encoding ASCII
                                $HealthAndStatus.Server.SystemInfo | Out-File "$PathToTransferDirectory\Server Info" -Encoding ASCII
                                $HealthAndStatus.Server.LogList    | Out-File "$PathToTransferDirectory\Server Logs" -Encoding ASCII


                                Write-Verbose "Modifying all Transfer Files to remove the uunwanted ANSI Color Codes."
                                
                                $Deliverables | ForEach-Object -Process {

                                        (Get-Content "$PathToTransferDirectory\$_")  `
                                                -replace $ANSIColorCodes, '' | 
                                                        Out-File "$PathToTransferDirectory\$_" -Encoding ASCII 

                                }

                        }

                Exit-PSSession

        }
        
        end {
                Write-Verbose "The Health and Status Custom PSObject Has now beem fully unpacked into the Transfer Directory."    
        }

}
Export-ModuleMember -Function Write-ToTransferDirectory




function Compare-LastWriteTimes {
<#
.SYNOPSIS
        Initiates a remote session between endpoints, and compares the Last Time any changes were Written into the Server's Transfer Directory
        with the Last Time any were Written into a File, given by a specified Path, that is meant to represent a System Activity Probe monitoring 
        system health and status.

.DESCRIPTION
        After checking whether it is either the Transfer Directory Files or the Specfied System Activity Indicator File, we return a string 
        simply stating which is the  Most Recently Modified File  out of the two comprators.  
                i.e.   Activity Probe   or   Transfer Directory

.PARAMETER UserName
        The Client's User Account Name belonging to the User who may authenticate to the Server.

.PARAMETER HostName
        The "Computer Name" assigned to the Server.

.PARAMETER PathToTransferDirectory
        The Path to a File Transfer Directory located on the Server. 

.PARAMETER PathToFileDemonstratingSystemActivity
        The Path to some arbtrary file located on the Server determined to demonstrate indications of verifiable system activty, 
        applicable to both internal and external systems resoectively.

.EXAMPLE
        Compare-LastWriteTimes `
                UserName = My_Name `
                HostName = Server_Name `
                PathToTransferDirectory = Path\to\directory
                PathToFileDemonstratingSystemActivity = Path\to\file
                
.EXAMPLE
        $params = @{
                UserName = My_Name;
                HostName = Server_Name;
                PathToTransferDirectory= Path\to\directory
                PathToFileDemonstratingSystemActivity = Path\to\file
        }

        Compare-LastWriteTimes @params -Verbose

.EXAMPLE
        $params = @{
                UserName = My_Name;
                HostName = Server_Name;
                PathToTransferDirectory= Path\to\directory
                PathToFileDemonstratingSystemActivity = Path\to\file
        }
        
        $params.Add( MostRecentlyModified, (Compare-LastWriteTimes @params) )

        Update-HealthAndStatus @params
#>

        [CmdletBinding()]
        
        param (
        
                [Parameter(Mandatory=$true)]
                [string]
                $UserName,

                [Parameter(Mandatory=$true)]
                [string]
                $HostName,

                [Parameter(Mandatory=$true)]
                [string]
                $PathToTransferDirectory,
                
                [Parameter(Mandatory=$true)]
                [string]
                $PathToFileDemonstratingSystemActivity
        
        )
        
        begin {

                Write-Verbose "The path to the Server's Transfer Directory is:  $PathToTransferDirectory."
                Write-Verbose "Initiating Session over SSH to  $UserName@$HostName."

                $Session = New-PSSession -UserName $UserName -HostName $HostName 

        }
        
        process {

                Write-Verbose ("Invoking a Command to Compare the Last Write Times between the Files in the Transfer Directory " +
                "and the specified File Demonstrating System Activity returning a string naming whichever that was observed to be " +
                "the Most Recently Modified of the two.")

                Invoke-Command `
                        -Session $Session `
                        -ArgumentList $PathToTransferDirectory,$PathToFileDemonstratingSystemActivity `
                        -OutVariable MostRecentlyModified `
                        -ScriptBlock {

                                Param($PathToTransferDirectory,$PathToFileDemonstratingSystemActivity)
                                
                                $LastTransferDirectoryWriteTime = 
                                        (Get-ChildItem "$PathToTransferDirectory")[0].LastWriteTime
                                

                                $LastActivityWriteTime = 
                                        (Get-ChildItem "$PathToFileDemonstratingSystemActivity").LastWriteTime
                        

                                (
                                        ($LastTransferDirectoryWriteTime -lt $LastActivityWriteTime) `
                                        ? `
                                        "Activity Probe" `
                                        : `
                                        "Transfer Directory" `
                                )

                        }

                Exit-PSSession  
        }
                
        end {           
                Write-Verbose "The Most Recently Modified File is:  $MostRecentlyModified."
        }

}
Export-ModuleMember -Function Compare-LastWriteTimes




function Update-HealthAndStatus {
<#
.SYNOPSIS
        Initiates a remote session between endpoints, and returns a Status Message describing the Health and Status
        of a specified system of intereest which is then Written to the Transfer Directory as a System Activity Report 
        used to Update the existing collection Health and Status Reports already present there.
        
.DESCRIPTION
        Generates a Status Message describing the current state of Activity within a System based on
        a given input source revealing the Most Recently Modified betweem the Files in the Transfer Directory,
        and an "Activity Probe" File produced autmotically by the System that indicates operational health and status.
        
        If it is determined that the Transfer Directory Files were the Most Recently Modified of the two, then the Message
        generated is:

                "As of $(Get-Date), there has been NO additional activity observed."

        Otherwise, it's the Activity Probe which was determined to be the Most Recently Modified, so the Message
        generated is:

                "Recent System Activity observed on $(Get-Date)"

        After the Message is generated, it is written to a new "System Activity" Health and Status file in the Transfer Directory
        if one does not already exist, or is simply Appended to an existing one, should one be found. 

.PARAMETER MostRecentlyModified
        A string indicating whether it was the Activity Probe File or one of those the Transfer Directory which has been  
        Modified Most Recently.

.PARAMETER UserName
        The Client's User Account Name belonging to the User who may authenticate to the Server.

.PARAMETER HostName
        The "Computer Name" assigned to the Server.

.PARAMETER PathToTransferDirectory
        The Path to a File Transfer Directory located on the Server. 

.PARAMETER PathToFileDemonstratingSystemActivity
        The Path to some arbtrary file located on the Server determined to demonstrate indications of verifiable system activty, 
        applicable to both internal and external systems respectively.

.EXAMPLE
        $params = @{
                UserName = My_Name;
                HostName = Server_Name;
                PathToTransferDirectory= Path\to\directory
                PathToFileDemonstratingSystemActivity = Path\to\file
        }
        
        $params.Add( MostRecentlyModified, (Compare-LastWriteTimes @params) )

        Update-HealthAndStatus @params
#>

        [CmdletBinding()]

        param (
                
                [Parameter(Mandatory=$true)]
                [string]
                $MostRecentlyModified,
        
                [Parameter(Mandatory=$true)]
                [string]
                $UserName,

                [Parameter(Mandatory=$true)]
                [string]
                $HostName,

                [Parameter(Mandatory=$true)]
                [string]
                $PathToTransferDirectory,
                
                [Parameter(Mandatory=$true)]
                [string]
                $PathToFileDemonstratingSystemActivity
        
        )
        
        begin {

                Write-Verbose ("The most recently modified between the Transfer Directory Files and the File " + 
                "Demonstrating System Activity:  $MostRecentlyModified.")

                Write-Verbose "The path to the Server's Transfer Directory is:  $PathToTransferDirectory."
                Write-Verbose "The path to the File Demonstrting System Activity is:  $PathToFileDemonstratingSystemActivity."


                Write-Verbose "Initiating Session over SSH to  $UserName@$HostName."

                $Session = New-PSSession -UserName $UserName -HostName $HostName 


                Write-Verbose "Generating a Message describing the Health and Status of the File Demonstrating System Activity"

                $ActivityHealthAndStatus = 
                
                        ($MostRecentlyModified -match 'Transfer Directory') `
                        ? `
                        "As of $(Get-Date), there has been NO additional activity observed." `
                        : `
                        "Recent System Activity observed on $(Get-Date)"

        }
        
        process {
                
                Write-Verbose ("Invoking a Command to Append the Generated System Activity Health and Status message " +
                "to the other Health and Status Reports currently residing within the Transfer Directory.")

                Invoke-Command `
                        -Session $Session `
                        -ArgumentList $PathToTransferDirectory, $ActivityHealthAndStatus `
                        -ScriptBlock {
                                Param($PathToTransferDirectory, $ActivityHealthAndStatus)
                                $PathToActivityFile = "$PathToTransferDirectory\System Activity"
                                
                                $ActivityHealthAndStatus | Out-File $PathToActivityFile -Append
                        }

                Exit-PSSession
        }
        
        end {
                Write-Verbose "The Update to Health and Status has now been written to the Transfer Directory."
        }
}
Export-ModuleMember -Function Update-HealthAndStatus