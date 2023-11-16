<#
.SYNOPSIS
By default script is using -whatif flag to allow for validation of actions prior to running.
Remove -whatif flag from second call of Get-NetworkAdapter to enact changes

The purpose of this script is to move all VM's from one VLAN to another using details outlined in a CSV

.Description
Create an Array of items by importing a CSV defined by $csvFilePath variable then loop through each entry, log every action, and log an error if an action failed.

#    CSV should be formatted with the following headers:
#    vcenter,sourceVLAN,targetVLAN,targetCluster
#
#    Example of data in the noted columns:
#    vcentername.contoso.com,vlan123,vlan354,Cluster_ENV_Name

The script will do the following for each line in the CSV.

    1. Call the change-vlan function
    2. Connect to vCenter
    3. Query all VM's that reside in the sourceVLAN & targetCluster 
    4. Loop through all VM's from the query
    5. Get the network adapter or adapters of each VM that reside in the sourceVLAN
    6. Update the VM network adapter to be on the targetVLAN and connected
    7. Disconnect from vCenter at the end of the list of VM's

.OUTPUTS
When the script is run a logfile in the same path the CSV file is picked up from.

.NOTES
  Version:        1.0
  Author:         Dan Menafro
  Creation Date:  10/4/2023
  Last Changed:   10/4/2023

  Purpose/Change:    

#>

#############
# Functions #
#############

    # Change vLAN
    Function change-vlan {
    [CmdletBinding()]
    param (
        [Parameter(Position=0,Mandatory=$true)]
        [string]$sourceVLAN,
        
        [Parameter(Position=1,Mandatory=$true)]
        [string]$targetVLAN,
        
        [Parameter(Position=2,Mandatory=$true)]
        [string]$clustername,

        [Parameter(Position=3,Mandatory=$true)]
        [string]$vCenter,

        [Parameter(Position=4,Mandatory=$true)]
        [string]$LogFile
        ) # End defining parameters

        #
        Try {
                    
            # Connect to vcenter
            $vConnect = Connect-VIServer -server $vCenter -Credential $Credential -Force -ErrorAction Stop

            # Log connection to vcenter
            update-log -type "Action" -data "connected to $vcenter" -LogFile $LogFile

            # Search for VM's in $source VLAN on $clustername
            Try {

                # Query vcenter for VM's in specified vlan
                $VMList = Get-VDPortgroup -ErrorAction Stop `
                    | Where-Object name -ceq $sourceVLAN `
                    | get-vm `
                    | Select-Object name,PowerState,@{name="VMHostName"; Expression={$_.VMHost.name}},@{name="ClusterName"; expression={$_.VMHost.Parent.Name}} `
                    | Where-Object Clustername -ceq $clustername
                
                # Loop through VM's and update their vlan
                foreach($VM in $VMList){
                    
                    # Avoid messing with a Zerto VM
                    if($VM.name -like "*Z-VRA*"){

                    write-host "Skipped" + $VM.name + " because it's a zerto system"
                    $data = $VM.name + ",because it's a zerto system," + $targetVLAN + "," + $vcenter + "," + $sourceVLAN + "," + $targetVLAN + "," + $clustername

                    update-log -type "Skipped" -data $data

                    } else {

                        # Triple checking the cluster name is correct just in case
                        If($vm.ClusterName -eq $Clustername){

                            # Perform VLAN change
                            Try {
                                
                                # Update vlan
                                #write-host "Moving" $VM.name "from" $sourceVLAN "to" $targetVLAN
                                                               
                                $NetworkAdapters = Get-NetworkAdapter -vm $vm.Name | Where-Object NetworkName -ceq $sourcevlan
                                
                                foreach($NIC in $NetworkAdapters){
                                    
                                    Try {                                  
                                        

                                        Get-NetworkAdapter -vm $vm.name -name $NIC.name | Set-NetworkAdapter -Connected $true -NetworkName $targetVLAN -Confirm:$false -WhatIf
                                        
                                        # Log change
                                        $data = "Updated network adapter " + $nic.name + " to " + $targetVLAN + " on " + $VM.name + "," + $vcenter + "," + $sourceVLAN + "," + $targetVLAN + "," + $clustername

                                        update-log -LogFile $LogFile -Type "Action" -data $data

                                    } catch {
                                        
                                        # Populate error
                                        $data = $Error[0].exception + "," + $VM.name + "," + $vcenter + "," + $sourceVLAN + "," + $targetVLAN + "," + $clustername

                                        # Error encountered during vlan update
                                        # Send data to log
                                        update-log -LogFile $LogFile -type "Error" -data $data

                                    } # End trying to move to different VLAN

                                } # End looping through all NIC's on the VM


                            } catch {
                                
                                # Populate error
                                $data = $Error[0].exception + "," + $VM.name + "," + $vcenter + "," + $sourceVLAN + "," + $targetVLAN + "," + $clustername

                                # Error encountered during vlan update
                                # Send data to log
                                update-log -LogFile $LogFile -type "Error" -data $data

                            } # End vlan changing
                        } # Closing cluster name validation
                    } # Closing Zerto VRA check
                } # End looping through VM's to change
            
            } catch {
                # Populate error
                $data = $Error[0].exception + "," + $VM.name + "," + $vcenter + "," + $sourceVLAN + "," + $targetVLAN + "," + $clustername

                # Error encountered during query
                # Send data to log
                update-log -LogFile $LogFile -type "Error" -data $data

            } # End query of vcenter
        
        } catch {
            
            # Error encountered connecting to vcenter

            # Populate error
            $data = $Error[0].exception + "," + $VM.name + "," + $vcenter + "," + $sourceVLAN + "," + $targetVLAN + "," + $clustername

            # Send data to log
            update-log -LogFile $LogFile -type "Error" -data $data

        } Finally {
            
            if($vConnect){

                # Disconnect from vCenter
                Disconnect-VIServer -Server $vCenter -Confirm:$false -Force
                
                # Clear the connection variable just in case there's stale data from a previous connection
                Clear-Variable vConnect

            } # End validating there was actually a connection to vCenter to disconnect

        } # End coneccting to vCenter
    } # End change-vlan function

    # Prompt the user to select a file for the script to use as a list
    Function Get-FileName($initialDirectory){
        [System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms") | Out-Null

        $OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog
        $OpenFileDialog.InitialDirectory = $initialDirectory
        $OpenFileDialog.Filter = "CSV (*.csv) | *.csv"
        $OpenFileDialog.Title = "VLAN move list"
        $OpenFileDialog.ShowDialog() | Out-Null
        $OpenFileDialog.FileName
    } # End Get-FileName function

    # Logging
    Function update-log{
        [CmdletBinding()]
        param (
            [Parameter(Position=0,Mandatory=$true)]
            [string]$type,
        
            [Parameter(Position=1,Mandatory=$true)]
            [string]$data,

            [Parameter(Position=1,Mandatory=$true)]
            [string]$LogFile
            ) # End defining parameters
        
        try {
            
            $timestamp = Get-Date -UFormat "[%Y.%m.%d]-%R"
            $logdata = "$timestamp,$type,$data"

            Add-Content -value $logdata -Path $LogFile
        
        }catch{
            write-host $error[0].exception -BackgroundColor Red
        } # End trying to update log

    }

#####################
# Module validation #
#####################

    # Check VMware PowerCLI is installed
    If(-not(get-module -ListAvailable VM*)){
        
        write-host "VMware PowerCLI is NOT installed." -BackgroundColor Red
        write-host "Exiting script. Go install VMware PowerCLI" -BackgroundColor Red
        
        Exit # Stop the script here since it doesn't meet the minimum requirements to proceed
       
    } # End check for VMware PowerCLI

#############
# Variables #
#############

    # Get date for use
    $date = get-date -UFormat "%Y.%m.%d"

    # Set log file name
    $LogFileName = "$date-vLAN_move.log"

    # Getting the desktop path of the user launching the script. Just as a default starting path
    $DesktopPath = [System.Environment]::GetFolderPath([System.Environment+SpecialFolder]::Desktop)
        
    $csvFilePath  = Get-FileName -initialDirectory $DesktopPath
    $csv      = @()
    $csv      = Import-Csv -Path $csvFilePath

    # Define where the log file should go (same folder as CSV file)
    $LogPath = $csvFilePath | Split-Path -Parent      
    
    # Combine log location with file name
    $LogFile = "$logpath\$logfilename"

    # Prompt for credentials
    $Credential = Get-Credential -Message "Please provide username in domain\user format"

###############
# Do the work #
###############

    # Loop through the list of VLANs

    foreach ($item in $csv){

       change-vlan -sourceVLAN $item.sourceVLAN -targetVLAN $item.targetVLAN -vCenter $item.vcenter -ClusterName $item.targetCluster -LogFile $LogFile
        
    } # End looping through VLANs

    write-host "Script stopped see log at: $logfile"
