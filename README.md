# GPO-Analyzer-Automation-Tool
Tool that can convert and upload GPOs automatically to Microsoft Endpoint Managers Group Policy Analytics tool.


.SYNOPSIS

    Enumerates GPOs in use in on premises AD, converts them to XML and then
    uploads them to GPO Analitics in MEMC


.DESCRIPTION

    This script automates the conversion and upload of GPOs to the GPO Analitics
    tool in MEMC for conversion to MEM policies

    To use this script, you will need:

    1. An Internet connection
    2. Line of site to a domain controller
    3. MEMC administrator account, or sufficest privlidges to upload via Graph
    4. The ability to install MS-Graph and AD PS modules if not already installed.
   

.EXAMPLE
    .\Upload-GPOReportsToMEM.ps1

    Requires and checks for MS graph and AD PS modules
    
    Will prompt for MEMC admin credentials to ingest GPOs.

.NOTES

    The purge feature removes all uploaded GPOs from GPO Analitics but will not
    delete any policies created from the migration of them.  

#>

<#
Revision history
    1.0.0  - Initial release  ActiveDirectory
#>