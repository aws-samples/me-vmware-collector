#Disabling SSL check
Set-PowerCLIConfiguration -InvalidCertificateAction Ignore -Confirm:$false 
#disabling the VMWare Customer Experience Improvement Program
Set-PowerCLIConfiguration -Scope User -ParticipateInCEIP  $false
Import-Module ImportExcel
# Create a new Excel file

Write-Host "Connecting to vCenter Server..."
$ipAddress = Read-Host -Prompt "Enter the IP address"
$username = Read-Host -Prompt "Enter the username"
$password = Read-Host -Prompt "Enter the password" 
$userInput = Read-Host "Enter the number of collection days"
#$filterVMs = Read-Host "Do you want to consider Powered ON only VMs(Y/N)"
$collectionDays = [int]::Parse($userInput)

Connect-VIServer $ipAddress -User $username -Password $password
$allvms = @()
$dailyStats = @()
$allhosts = @()
Write-Host "Retrieving all VMs from vCenter..."
$vms= Get-Vm | Where {$_.PowerState -eq "PoweredOn"}

$hosts = Get-VMHost

foreach ($esxi in $hosts) {
    Write-Host "Collecting host information for $($esxi.Name)..."

    $hoststat = New-Object PSObject -Property @{
        "HostName"      = $esxi.Name
        "NumCPUs"       = $esxi.NumCpu
        "MemoryMB"      = $esxi.MemoryTotalMB
        "ProcessorType" = $esxi.ProcessorType
    }
	$allhosts += $hoststat	
}

foreach ($vm in $vms) {
    Write-Host "Processing VM:" $vm.Name
    Write-Host "Collecting basic information..."
    $vmstat = New-Object PSObject -Property @{
        "ServerName" = $vm.Name  
        "OperatingSystem" = $vm.Guest.OSFullName
        "NumCPUs" = $vm.NumCpu
        "MemoryVM" = $vm.MemoryMB 
        "StorageGB" = 0 
        "HostName"= $vm.VMhost
    }

    Write-Host "Calculating total storage..."
    $disks = Get-HardDisk -VM $vm
    foreach ($disk in $disks) {
        $vmstat.StorageGB += [math]::Round($disk.CapacityGB, 2)  # Add disk size in GB with 2 decimals
    }

    for ($i = 0; $i -lt $collectionDays; $i++) {
        $dayStart = (Get-Date).AddDays(-$i).Date  # Start of the day
        $dayEnd = $dayStart.AddDays(1).AddSeconds(-1)  # End of the day
		if ($i -le 1) {
			$intervalSecs = 300  # 5 minutes
		}
		elseif ($i -ge 2 -and $i -le 7) {
			$intervalSecs = 1800  # 30 minutes
		}
		elseif ($i -gt 7 -and $i -le 30) {
			$intervalSecs = 7200  # 2 hours
		}
		else {
			$intervalSecs = 86400  # 1 day
		}
        Write-Host "Retrieving peak CPU and memory usage for" $dayStart
        # Get peak CPU and memory usage for the specific day
        $statcpu = Get-Stat -Entity $vm -Start $dayStart -Finish $dayEnd -MaxSamples 20000 -Stat cpu.usage.average -IntervalSecs $intervalSecs
        $statmem = Get-Stat -Entity $vm -Start $dayStart -Finish $dayEnd -MaxSamples 20000 -Stat mem.consumed.average -IntervalSecs $intervalSecs
        # Calculate peak and average values
        $CpuPeakValue = ($statcpu | Measure-Object -Property Value -Maximum | Select-Object -ExpandProperty Maximum)
        $CpuAvgValue= ($statcpu | Measure-Object -Property Value -Average | Select-Object -ExpandProperty Average)
        $MemoryPeakValue = ($statmem | Measure-Object -Property Value -Maximum | Select-Object -ExpandProperty Maximum)
        $MemoryAvgValue = ($statmem | Measure-Object -Property Value -Average | Select-Object -ExpandProperty Average)
        $cpuPeak = [math]::Round($CpuPeakValue, 2)
		$memPeak = [math]::Round($MemoryPeakValue / (1024 * $vm.MemoryMB), 2)
		$cpuAvg = [math]::Round($CpuAvgValue, 2)
		$memAvg = [math]::Round($MemoryAvgValue / (1024 * $vm.MemoryMB), 2)


        $dailyStat = New-Object PSObject -Property @{
            "ServerName" = $vm.Name 
            "Date" = $dayStart
            "CpuPeak" = $cpuPeak
            "CpuAvg" = $CpuAvg
            "MemPeak" = $memPeak
            "MemAvg" = $memAvg
        }
        $dailyStats += $dailyStat
    }
    $allvms += $vmstat
}

$output = "VMWARE_Inventory_And_Usage_Workbook_" + (Get-Date).ToString("yyyy-MM-dd")+".xlsx"

# Exporting provisioning and utilization metrics to output inventory file
try {
	
	$src_provisioning = $allvms
	$vms_provisioning = $null
	$phs_provisioning = $null
	$ast_ownership = $null
	$vms_utilization = $null
	$i=1
	$name_uid         = @{}
	
    $allhosts | ForEach-Object {
		$name = $_."HostName"
		$ph_data = @(
			[PSCustomObject]@{
				"Unique Identifier"         = $(New-Guid)
				"Human Name"                = $name
                "pCpu Cores" 				= $_."NumCPUs"
                "Memory MB" 				= $_."MemoryMB"
                "Total Storage Size GB" = ""
                "Cpu String" 				= $_."ProcessorType"
                "Operating System" = ""
                "Database Type" = ""
                "Address" = ""
                "Remote Storage Size GB" = ""
                "Remote Storage Type" = ""
                "Local Storage Size GB" = ""
                "Local Storage Type" = ""
                "Location" = ""
                "Make" = ""
                "Model" = ""
				}
			)
		$phs_provisioning  += $ph_data
		} 
    $phs_provisioning | Export-Excel -Path $output -AutoSize -WorksheetName "Physical Provisioning"

	$src_provisioning | ForEach-Object {
		# Virtual Provisioning
		$uid                            = $(New-Guid)
		$name                           = $_."ServerName"
		$vm_data = @(
		  [PSCustomObject]@{
			"Unique Identifier"         = $uid
			"Human Name"                = $name
			"vCpu Cores"                = $_."NumCPUs"
			"Memory MB"                 = $_."MemoryVM"
			"Total Storage Size GB"     = $_."StorageGB"
			"Operating System"          = $_."OperatingSystem"
            "Database Type" = ""
			"Hypervisor Name"           = $_."HostName"
            "Address" = ""
			"Remote Storage Size GB"  ="" 
            "Remote Storage Type" = ""
            "Local Storage Size GB" = $_."StorageGB"
            "Local Storage Type" = ""
			}
		)
		$vms_provisioning  = $vms_provisioning + $vm_data

		# Asset Ownership
		$ao_data = @(
		  [PSCustomObject]@{
			"Unique Identifier"         = $uid
			"Human Name"                = $name
            "Environment" = ""
            "Application"= ""
            "SLA" = ""
            "Department" = ""
            "Line of Business" = ""
            "In Scope" = ""
			}
		)
		$ast_ownership += $ao_data
		$name_uid[$name]   = $uid
		$i++
	}
	$vms_provisioning | Export-Excel -Path $output -AutoSize -WorksheetName "Virtual Provisioning"
	$ast_ownership   | Export-Excel -Path $output -AutoSize -WorksheetName "Asset Ownership"
}

catch {
	Write-Error "[ERROR]provisioning data conversion failed!"
	Exit
} 

try {
	# Utilization
	$src_utilization  = $dailyStats
	$src_utilization | ForEach-Object {
		$name                                = $_."ServerName"
		$date                                = [DateTime]::Parse($_.Date, [cultureinfo]::GetCultureInfo('en-us'))
		$ut_data = @(
		  [PSCustomObject]@{
			"Unique Identifier"	             = $name_uid[$name]
			"Human Name"                     = $name
			"Cpu Utilization Peak (P95)"     = $_.CpuPeak / 100
			"Memory Utilization Peak (P95)"  = $_.MemPeak
			"Storage Utilization Peak (P95)" = 1
			"Cpu Utilization Avg (P95)"      = $_.CpuAvg / 100
			"Memory Utilization Avg (P95)"   = $_.MemAvg
			"Storage Utilization Avg (P95)"  = 1
			"Time On Percentage"             = 1
			"Time In-Use Percentage"         = 1
			"Time Stamp Start"               = $date.ToString("yyyy-MM-ddT00:00:00.0000000")
			"Time Stamp End"                 = $date.AddDays(1).ToString("yyyy-MM-ddT00:00:00.0000000")
			}
		)
		$vms_utilization  = $vms_utilization + $ut_data
	}

	$vms_utilization  | Export-Excel -Path $output -AutoSize -WorksheetName "Utilization"
}
catch {
	Write-Error "[ERROR]utilization data conversion failed!"
	Exit
}