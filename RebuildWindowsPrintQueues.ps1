############################################################################################
#Launch an elevated PowerShell widow

#Get the ID and security principal of the current user account
 $myWindowsID=[System.Security.Principal.WindowsIdentity]::GetCurrent()
 $myWindowsPrincipal=new-object System.Security.Principal.WindowsPrincipal($myWindowsID)
  
 #Get the security principal for the Administrator role
 $adminRole=[System.Security.Principal.WindowsBuiltInRole]::Administrator
  
 #Check to see I am currently running "as Administrator"
 if ($myWindowsPrincipal.IsInRole($adminRole))
    {
    #I am running "as Administrator" - so change the title and background color to indicate this
    $Host.UI.RawUI.WindowTitle = $myInvocation.MyCommand.Definition + "(Elevated)"
    $Host.UI.RawUI.BackgroundColor = "Black"
    clear-host
    }
 else
    {
    #I am not running "as Administrator" - so relaunch as administrator
    
    #Create a new process object that starts PowerShell
    $newProcess = new-object System.Diagnostics.ProcessStartInfo "PowerShell";
    
    #Specify the current script path and name as a parameter
    $newProcess.Arguments = $myInvocation.MyCommand.Definition;
    
    #Indicate that the process should be elevated
    $newProcess.Verb = "runas";
    
    #Start the new process
    [System.Diagnostics.Process]::Start($newProcess);
        
    
    #Exit from the current, unelevated, process
    exit
    }

############################################################################################  
#Adjust window size
$h = get-host
$win = $h.ui.rawui.windowsize
$win.width  = 100 
$win.Height = 30
$h.ui.rawui.set_windowsize($win)

#Script
############################################################################################
#Specifying Windows Print Servers and available drivers
$PrintServers = ("tec-v-prntsrv01", "tec-v-prntsrv02", "tec-v-prntsrv03", "tec-v-prntsrv04")
$PrintServerDrivers = Get-PrinterDriver -ComputerName tec-v-prntsrv01 -PrinterEnvironment "Windows x64" | sort Name | select -ExpandProperty name -ErrorAction SilentlyContinue

############################################################################################
#Load the .Net Assembly for my PopUp boxes
[void][System.Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic')

############################################################################################
#Prompt for the Printer Name
$PrinterName = [Microsoft.VisualBasic.Interaction]::InputBox('Enter the Host Name of the Printer', 'Create New Print Queues')
Write-Host "Gathering configuration of $PrinterName..." -ForegroundColor Yellow

############################################################################################
#Prompt for Printer IP Address
$IPAddress = [Microsoft.VisualBasic.Interaction]::InputBox('Enter the IP Address of the Printer', 'Create New Print Queues')

############################################################################################
#Prompt for Print Driver
$objForm = New-Object System.Windows.Forms.Form 
$objForm.Text = "Select a Printer Driver"
$objForm.Size = New-Object System.Drawing.Size(390,500) 
$objForm.StartPosition = "CenterScreen"
$OKButton = New-Object System.Windows.Forms.Button
$OKButton.Location = New-Object System.Drawing.Size(90,12)
$OKButton.Size = New-Object System.Drawing.Size(75,23)
$OKButton.Text = "OK"
$CancelButton = New-Object System.Windows.Forms.Button
$CancelButton.Location = New-Object System.Drawing.Size(200,12)
$CancelButton.Size = New-Object System.Drawing.Size(75,23)
$CancelButton.Text = "Cancel"
$OKButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
$objForm.Controls.Add($OKButton)
$objForm.AcceptButton = $OKButton
$CancelButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
$objForm.Controls.Add($CancelButton)
$objForm.CancelButton = $CancelButton
$objListBox = New-Object System.Windows.Forms.ListBox 
$objListBox.Location = New-Object System.Drawing.Size(55,50) 
$objListBox.Size = New-Object System.Drawing.Size(260,20) 
$objListBox.Height = 400

foreach($PrintServerDriver in $PrintServerDrivers){
        $objListBox.Items.Add($PrintServerDriver) 
}
$objForm.Controls.Add($objListBox)
cls
$objForm.Topmost = $True
cls
$result = $objForm.ShowDialog() 
if ($result -eq [System.Windows.Forms.DialogResult]::OK -and $objListBox.SelectedIndex -ge 0){
    cls
    $PrintDriver = $objListBox.Text
}

############################################################################################
#Prompt for adding text to Printer Location Box in Printer Properties
$Location = [Microsoft.VisualBasic.Interaction]::InputBox('Enter a description of where the printer is located', 'Rebuild Print Queues')

############################################################################################
#Prompt to Verify Printer Info was entered correctly
$a = new-object -comobject wscript.shell 
$intAnswer = $a.popup("Click Yes to build Windows Print Queues`nClick No to cancel`n`n`n Printer Name: $PrinterName`n`n IP Address: $IPAddress`n`n Driver: $PrintDriver`n`n Location: $Location", ` 
0,"Verify Printer Settings",4)
If ($intAnswer -eq 6) { 

############################################################################################
#Actual script that creates the print queues

    foreach ($PrintServer in $PrintServers) {
        Write-Host "Checking if $PrinterName exists on $PrintServer"
        $Printer = Get-Printer -ComputerName $PrintServer -Name $PrinterName -ErrorAction SilentlyContinue
        If ($Printer.Name -eq $Null) {
        Write-Host "$PrinterName already does not exist on $Printserver" -ForegroundColor Red
        }
        Else {
        Remove-Printer -AsJob -ComputerName $PrintServer -Name $PrinterName -ErrorAction SilentlyContinue -Verbose
        }
        Write-Host "The print queue has been successfully deleted" -ForegroundColor Green
}

    foreach ($PrintServer in $PrintServers) {
        Write-Host "Checking if printer port $IPAddress exists on $PrintServer..."
        $Port = Get-Printerport -ComputerName $PrintServer -Name $IPAddress -ErrorAction SilentlyContinue
        If ($Port.Name -eq $Null) {
        Add-PrinterPort -AsJob -ComputerName $PrintServer -Name $IPAddress -PrinterHostAddress $IPAddress -ErrorAction SilentlyContinue -Verbose
        }
        Else {
        Write-Output "Printer port exists on $PrintServer. Moving on to building printer."
        }
        Write-Host "Checking if $PrinterName exists on $PrintServer..."
        $Printer = Get-Printer -ComputerName $PrintServer -Name $PrinterName -ErrorAction SilentlyContinue
        If ($Printer.Name -eq $Null) {
        Add-Printer -AsJob -ComputerName $PrintServer -Name $PrinterName -DriverName "$PrintDriver" -ShareName $PrinterName -PortName $IPAddress -Location $Location -Comment $IPAddress -Shared -Published -ErrorAction SilentlyContinue -Verbose
        }
        Else {
        Write-Output "$PrinterName already exists on $PrintServer."
        }
        Write-Host "The print queue is successfully built.`n`n" -ForegroundColor Green
    }

############################################################################################
#Notification that the script has completed. 
$a.popup("$PrinterName has been successfully built on all Windows Print Servers.")
Exit
}

 else { 
    $a.popup("Build Windows Print Queues has been cancelled")
    Exit 
} 