Splice Admin
============

Splice Admin is a remote Windows administration tool.  It allows you to retrieve information and interact with remote machines on your network.

### [Click here to download the latest .exe](https://github.com/R-Smith/Splice-Admin/releases/download/v2017.0101/SpliceAdm.exe)
### [Click here to download the source](https://github.com/R-Smith/Splice-Admin/archive/master.zip)

##### Notes
* There is no installer.  Just run the .exe.
* .NET 4.0 or greater is required on the machine running Splice Admin.
* Most features require WMI to be running on the target computer.
* Local administrative rights are required on the target computer.

### Changes for v2017.0101
If you haven't seen the bulk query feature yet, check it out.  You can easily search multiple computers (or even all computers) in your environment for installed applications, services, running processes, logged in users, and files or directories.
* Completely revamped the bulk query feature.  Rather than a simple 'match' or 'no match', searches now include detailed results.  For example, if searching for installed applications and your search phrase is 'adobe', you will get a list which includes the full title of every application with 'adobe' in the name, as well as the version number, for every computer in your target list.   When searching services, search results will include the startup type and the current status of the service.
* Bulk query:  The pre-defined target computer groups now exclude disabled computers.
* Bulk query:  New option to only include computers that have been active in the past 30 days.
* Bulk query:  Added the ability to search for running processes.
* Applications:  Remote applications were previously retrieved through the Remote Registry service.  The registry values are now retrieved through WMI which results in faster data retrieval.  The application will fall back to using the Remote Registry service if the target OS does not support the WMI registry provider.
* Bulk query:  Searching for applications now also goes through the WMI registry provider with a fallback to remote registry.  This results in faster application searches.
* Added: Tasks -> Windows Remote Assistance.  A shortcut to connect to the target computer using the Windows Remote Assistance tool.  This will only work if your environment is configured to allow unsolicited remote assistance.


Features
========
All features apply to remote Windows machines on your network.
* System Information
  * Get computer name, operating system, and uptime.
  * Get PC manufacturer, model, CPU and memory information.
  * Show connected USB devices.
  * Retrieve detailed settings for all connected network adapters.
  * Retrieve and edit system ODBC DSNs.
* Processes
  * Show all running processes along with their PID, owner, and path.
  * Color coding helps you spot user processes versus system processes.
  * Terminate any process (accessed via right-click).
* Services
  * Show all runing services along with their current status and startup type.
  * Color coding helps you spot different startup type and statuses.
  * Start and stop any service (accessed via right-click).
* Storage
  * Display all connected storage devices and their capacity information.
* Applications
  * Show all installed applications.
* Updates
  * Display all installed updates along with their installation date (if available) and KB article links.
  * Show if there is a pending required reboot.
  * Display update configuration (automatic or manual), the date and time updates were last checked (XP - 8.1), and the date and time updates were last installed (XP - 8.1).
  * Uninstall any update (accessed via right-click, Vista/2008 or higher).
* Users
  * Show all logged on users along with their login date/time, IP address (if available), and session ID.
  * Send a message to **any** user and logoff any user (accessed via right-click).
  * Retrieve basic details about the user from Active Directory (accessed via right-click).
  * Show a history of all logons and logoffs along with their date/time, IP address (if available), and whether it's an interactive or RDP connection.  Vista/2008 or higher.  Logon history depends on event logs so if the target isn't logging this information, you won't see it.
* Misc
  * Reboot / shutdown target computer.
  * Force a group policy update.
  * Execute any command or script on the target computer.  You do not get the output, the command is simply executed.
  * Launch various MMCs targeting the selected computer (Event viewer, computer management, services).


Screenshots
===========
##### Overview
![Overview](https://github.com/R-Smith/supporting-docs/raw/master/Splice-Admin/spliceadm-overview.png?raw=true "Overview")

##### Processes
![Processes](https://github.com/R-Smith/supporting-docs/raw/master/Splice-Admin/spliceadm-processes.png?raw=true "Processes")

##### Storage
![Storage](https://github.com/R-Smith/supporting-docs/raw/master/Splice-Admin/spliceadm-storage.png?raw=true "Storage")
