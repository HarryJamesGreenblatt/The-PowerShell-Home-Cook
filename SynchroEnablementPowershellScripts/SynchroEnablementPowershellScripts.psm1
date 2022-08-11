function Get-HealthAndStatus {

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

                $Session = New-PSSession -UserName $UserName -HostName $HostName

                $ClientHealthAndStatus = @{

                        SystemInfo = Get-ComputerInfo;

                        LogList = Get-WinEvent -ListLog * | 
                                                Where-Object RecordCount -gt 0 | 
                                                         Sort-Object RecordCount -Descending  
                }
                    
        }
        

        process {
                
                Invoke-Command `
                        -Session $Session `
                        -OutVariable ServerHealthAndStatus `
                        -ScriptBlock {
                    
                                @{
                                        SystemInfo = Get-ComputerInfo;
                                        
                                        LogList = Get-WinEvent -ListLog * | 
                                                        Where-Object RecordCount -gt 0 | 
                                                                        Sort-Object recordcount -Descending
                                }

                        }

                Exit-PSSession

        }
        
        end {

                [PSCustomObject]@{
                        Client = $ClientHealthAndStatus;
                        Server = $ServerHealthAndStatus
                }
                
        }
}
Export-ModuleMember -Function Get-HealthAndStatus                     





function Write-ToTransferDirectory {
        
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

                $Session = New-PSSession -UserName $UserName -HostName $HostName

                $Deliverables = @(
                        "Client Info",
                        "Client Logs",
                        "Server Info",
                        "Server Logs"
                )   
        }
        
        process {

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


                                If( -not (Test-Path $PathToTransferDirectory) ){
                                        New-Item -ItemType Directory $PathToTransferDirectory
                                }


                                $HealthAndStatus.Client.SystemInfo | Out-File "$PathToTransferDirectory\Client Info" -Encoding ASCII
                                $HealthAndStatus.Client.LogList    | Out-File "$PathToTransferDirectory\Client Logs" -Encoding ASCII
                                $HealthAndStatus.Server.SystemInfo | Out-File "$PathToTransferDirectory\Server Info" -Encoding ASCII
                                $HealthAndStatus.Server.LogList    | Out-File "$PathToTransferDirectory\Server Logs" -Encoding ASCII


                                $Deliverables | ForEach-Object -Process {

                                        (Get-Content "$PathToTransferDirectory\$_")  `
                                                -replace $ANSIColorCodes, '' | 
                                                        Out-File "$PathToTransferDirectory\$_" -Encoding ASCII 

                                }

                        }

                        
                Exit-PSSession

        }
        
        end {
                
        }

}
Export-ModuleMember -Function Write-ToTransferDirectory




function Compare-LastWriteTimes {

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
                $Session = New-PSSession -UserName $UserName -HostName $HostName                
        }
        
        process {
                
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
        }

}
Export-ModuleMember -Function Compare-LastWriteTimes




function Update-HealthAndStatus {

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
                $Session = New-PSSession -UserName $UserName -HostName $HostName 

                $ActivityHealthAndStatus = 
                
                        ($MostRecentlyModified -match 'Transfer Directory') `
                        ? `
                        "As of $(Get-Date), there has been NO additional activity observed." `
                        : `
                        "Recent System Activity observed on $(Get-Date)"

        }
        
        process {
                
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
        }
}
Export-ModuleMember -Function Update-HealthAndStatus