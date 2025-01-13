<<<<<<< HEAD
## My Project

TODO: Fill this README out!

Be sure to:

* Change the title in this README
* Edit your repository description on GitHub

## Security

See [CONTRIBUTING](CONTRIBUTING.md#security-issue-notifications) for more information.

## License

This library is licensed under the MIT-0 License. See the LICENSE file.

=======
# VMware Data Collector

## Description
This PowerShell script collects detailed inventory and performance metrics from a VMware vCenter environment. It gathers provisoning information about virtual machines, host systems, and their utilization percentages over a specified period, then exports the data to an Excel workbook.

## Prerequisites
- PowerCLI module installed
- ImportExcel PowerShell module installed 
- Access to vCenter Server
- Appropriate permissions to query VM and host statistics

## Features
- Collects host information including CPU, memory, and processor details
- Gathers VM specifications including CPU, memory, and storage allocation
- Captures performance metrics with different sampling intervals based on historical depth:
  - Last 2 days: 5-minute intervals
  - 2-7 days: 30-minute intervals
  - 8-30 days: 2-hour intervals
  - Beyond 30 days: Daily intervals
- Exports data to Excel with multiple worksheets:
  - Physical Provisioning
  - Virtual Provisioning
  - Asset Ownership
  - Utilization

## Usage
1. Run the script in PowerShell
2. Enter the requested information when prompted:
   - vCenter IP address
   - Username
   - Password
   - Number of collection days

## Output
The script generates an Excel file named "VMWARE_Inventory_And_Usage_Workbook_YYYY-MM-DD.xlsx" containing:
- Detailed host information
- VM configurations
- Resource utilization metrics
- Asset ownership details

## Notes
- VMware Customer Experience Improvement Program (CEIP) is disabled
- Only powered-on VMs are included in the collection by default
- Performance metrics include CPU and memory usage (peak and average values)

## Data Collection Details
### Host Information:
- Hostname
- CPU count
- Memory capacity
- Processor type

### VM Information:
- Server name
- Operating system
- CPU allocation
- Memory allocation
- Storage capacity
- Host assignment

### Performance Metrics:
- CPU peak and average usage
- Memory peak and average consumption

##Read More: 
https://aws.amazon.com/blogs/migration-and-modernization/accelerating-migration-evaluator-discovery-for-vmware-environment/
>>>>>>> d2d789c (Initial commit)
