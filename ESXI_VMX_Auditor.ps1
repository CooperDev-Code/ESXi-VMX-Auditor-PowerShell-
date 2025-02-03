# Script configuration
[CmdletBinding()]
param(
    # Add validation for vCenter parameter
    [Parameter(Mandatory=$true)]
    [ValidateNotNullOrEmpty()]
    [string]$vCenter,
    
    [Parameter(Mandatory=$true)]
    [string]$DSName,
    
    [Parameter()]
    [string]$VMXFolder = "C:\vmx",
    
    [Parameter()]
    [System.Management.Automation.PSCredential]
    [System.Management.Automation.Credential()]
    $Credential = (Get-Credential -Message "Enter vCenter credentials")
)

# Script Variables
$ErrorActionPreference = "Stop"
$ScriptPath = Split-Path -Parent $MyInvocation.MyCommand.Path
$ConnectionTimeout = 30  # Connection timeout in seconds
$VMXLogFile = Join-Path $VMXFolder "vms-on-$DSName-$(Get-Date -Format 'ddMMyy-hhmmss').csv"
$horLine = "----------------------------------------------------------------------------------------------------------------------------------------"

# Initialize PowerCLI
try {
    Set-PowerCLIConfiguration -Scope User -InvalidCertificateAction Warn -Confirm:$false
} catch {
    Write-Error "Failed to configure PowerCLI: $_"
    exit 1
}

# Ensure VMX folder exists
if (-not (Test-Path $VMXFolder)) {
    New-Item -ItemType Directory -Path $VMXFolder -Force | Out-Null
}

#Connect to vCenter Server
try {
    # Test if we can reach the vCenter server
    if (-not (Test-Connection -ComputerName $vCenter -Count 1 -Quiet)) {
        throw "Cannot reach vCenter server $vCenter"
    }
    
    # Check if we're already connected
    $existingConnection = $global:DefaultVIServer | Where-Object { $_.Name -eq $vCenter }
    if ($existingConnection) {
        Write-Verbose "Already connected to $vCenter"
    } else {
        $connection = Connect-VIServer -Server $vCenter -Credential $Credential -ErrorAction Stop
        if (-not $connection) {
            throw "Failed to establish connection to vCenter"
        }
        Write-Verbose "Successfully connected to $vCenter"
    }
} catch {
    Write-Error "vCenter Connection Error: $_"
    exit 1
}

clear 

#If datastore name is specified incorrectly by the user, terminate
try {$DSObj = Get-Datastore -name $DSName -ErrorAction Stop} 
catch {Write-Host "Invalid datastore name" ; exit}

#Get datastore view using id
$DSView = Get-View $DSObj.id

#Name is case-sensitive hence the need to retrieve the name even though specified by user
$DSName = $DSObj.name

#Fetch a list of folders and files present on the datastore
$searchSpec = New-Object VMware.Vim.HostDatastoreBrowserSearchSpec
$DSBrowser = get-view $DSView.browser
$RootPath = ("[" + $DSView.summary.Name + "]")
$searchRes = $DSBrowser.SearchDatastoreSubFolders($RootPath, $searchSpec)

#Object Counter
$s=0; 

#Get a list of virtual machines and templates residing on the datastore
$vms=(get-vm * -Datastore $DSObj).Name
$templates=(get-template * -Datastore $DSObj).Name 

#Write header to log file
("#,Type,VM_Name,VMX_Filename,VM_Folder,Name_Match?,Is_VM_Registered,ESXi_Host") | 
Out-File -FilePath $VMXLogFile -Append

#Write table header row to console
Write-Host "Browsing datastore $DSObj ...`n"
Write-Host $horLine
Write-Host ($colSetup -f "#", "Type", "VM Name", "VMX Filename", "Folder Path [$DSName]","Match?" , "Reg?", "ESXi Host")  -ForegroundColor white
Write-Host $horLine

#Recursively check every folder under the datastore's root for vmx files.
foreach ($folder in $searchRes)
{
    $type = $null      #Template or virtual machine?
    $VMXFile = $null   #Stores vmx/vmtx filename
    $registered = "No" #Is the virtual machine registered?
    $nameMatch = "No"  #Does the folder name match that of the virtual machine?
    $col = "Green"     #Default console color

    $DCName = $DSObj.Datacenter.Name
    $VMFolder = (($folder.FolderPath.Split("]").trimstart())[1]).trimend('/') 
    $VMXFile = ($folder.file | where {$_.Path -like "*.vmx" -or $_.Path -like "*.vmtx"}).Path  #vmtx is for templates
    $VMPath = ($DSName + "/" + $VMFolder)
    $fileToCopy = ("vmstore:/" + $DCName + "/" + $VMPath + "/" + $VMXFile)

    #Assuming vmx file exists ...
    if ($VMXFile -ne $null)
    {
        $s++

        #Extract VM name from the vmx file name. We will compare this to the value returned by displayName
        if ($VMXFile.contains(".vmx")){$prevVMName = $VMXFile.TrimEnd(".vmx"); $type="VM"} #Virtual Machine
        elseif ($VMXFile.contains(".vmtx")){$prevVMName = $VMXFile.TrimEnd(".vmtx"); $type="Template"} #Template

        #Copy vmx file to a local folder
        copy-DatastoreItem $fileToCopy $VMXFolder -ErrorAction SilentlyContinue

        #Extract the current virtual machine name from the VMX file as well as the host name
        Try
        {
            $owningVM = ((get-content -path ($VMXFolder + "/" + $VMXFile) -ErrorAction SilentlyContinue | 
            Where-Object {$_ -match "displayName"}).split(""""))[1]
            
            if ( $type.equals("VM")){$vmHost = (Get-VM -Name $owningVM -ErrorAction SilentlyContinue).vmhost}
            else {$vmHost = (Get-template -Name $owningVM -ErrorAction SilentlyContinue).vmhost}

            if ($vmHost -eq $null) {$vmHost="n/a"}
        }
        Catch 
        {
            $owningVM="Error retrieving ..."
            $vmHost="Error ..."
        } 

        #If the virtual machine specified in the VMX file is found in the list of virtual machines or templates, mark it as registered
        if (($vms -contains $owningVM) -or ($templates -contains $owningVM)) {$registered = "Yes"} else {$col="Red"}

        #Check folder name. Set $nameMatch to true if no conflict found
        if ($prevVMName.equals($owningVM) -and $prevVMName.equals($VMFolder)){$nameMatch="Yes"} else {$col="Red"};

        #Highlight unregistered virtual machines in cyan
        if ($registered.Equals("No")){$col="Cyan"}

        #Update Logfile
        ($s.ToString() + "," + $type + "," + $owningVM + "," + $VMXFile + "," + $VMFolder + "," + $nameMatch + "," + $registered + "," + $vmHost) | 
        Out-File -FilePath $VMXLogFile -Append 

        #Truncate strings if they do not fit the respective column width
        if ($owningVM.Length -ge 30) {$owningVM = (($owningVM)[0..26] -join "") + "..."}
        if ($VMXFile.Length -ge 30) {$VMXFile = (($VMXFile)[0..26] -join "") + "..."}
        if ($VMFolder.Length -ge 40) {$VMFolder = (($VMFolder)[0..36] -join "") + "..."}

        #Write to console
        write-host ($colSetup -f $s.ToString() , $type , $owningVM , $VMXFile, $VMFolder, $nameMatch, $registered, $vmHost) -ForegroundColor $col
    } 
}

Write-Host $horLine
# Cleanup section
try {
    Write-Verbose "Disconnecting from vCenter..."
    Disconnect-VIServer -Server $vCenter -Force -Confirm:$false
} catch {
    Write-Warning "Error during disconnect: $_"
}
