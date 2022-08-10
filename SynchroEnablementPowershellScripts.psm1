Function Get-HealthAndStatus {

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
                $TransferDirectoryPath

        )
        
        begin {

                $Session = New-PSSession -UserName $UserName -HostName $HostName

                $ClientHealthAndStatus = @{

                        SystemInfo = Get-ComputerInfo;

                        LogList = Get-WinEvent -ListLog * | 
                                                Where-Object RecordCount -gt 0 | 
                                                         Sort-Object recordcount -Descending  
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
        }
        
        end {
                [PSCustomObject]@{
                        Client = $ClientHealthAndStatus;
                        Server = $ServerHealthAndStatus
                }
        }
}
Export-ModuleMember -Function 'Get-HealthAndStatus'                     





Function Write-ToTransferDirectory {
        
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
                $TransferDirectoryPath,
                
                [string]
                $ANSIColorCodes = '\x1B\[([0-9]{1,3}((;[0-9]{1,3})*)?)?[m|K]'           
    
        )
        
        begin {

                $Deliverables = @("Client Info", "Client Logs", "Server Info", "Server Logs")   
                $Session = New-PSSession -UserName $UserName -HostName $HostName

        }
        
        process {
                Invoke-Command `
                        -Session $Session `
                        -ArgumentList $HealthAndStatus,$TransferDirectoryPath `
                        -OutVariable ServerHealthAndStatus `
                        -ScriptBlock {

                                Param([PSCustomObject]$HealthAndStatus, $TransferDirectoryPath)

                                $HealthAndStatus.Client.SystemInfo | Out-File "$TransferDirectoryPath\Client Info" -Encoding ASCII
                                $HealthAndStatus.Client.LogList    | Out-File "$TransferDirectoryPath\Client Logs" -Encoding ASCII
                                $HealthAndStatus.Server.SystemInfo | Out-File "$TransferDirectoryPath\Server Info" -Encoding ASCII
                                $HealthAndStatus.Server.LogList    | Out-File "$TransferDirectoryPath\Server Logs" -Encoding ASCII

                        }
        }
        
        end {
                Invoke-Command `
                        -Session $Session `
                        -ArgumentList $Deliverables,$TransferDirectoryPath,$ANSIColorCodes `
                        -OutVariable ServerHealthAndStatus `
                        -ScriptBlock {

                                Param([System.Object[]]$Deliverables, $TransferDirectoryPath, $ANSIColorCodes)

                                $Deliverables | ForEach-Object -Process {

                                        (Get-Content "$TransferDirectoryPath\$_")  `
                                                -replace $ANSIColorCodes, '' | 
                                                        Out-File "$TransferDirectoryPath\$_" -Encoding ASCII 

                                }
                        }
        }

}
Export-ModuleMember -Function 'Write-ToTransferDirectory'  