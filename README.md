# Fileshare Connection Test --- Visual Basic Script
A visual basic script to verify that the network connection to a file share is stable enough (fast and low latency) to open a MS Access database.

# How
Implementing this script in/as an Access database launching script attempts to prevent teleworking users from connecting to the databases unless they use Citrix. The completed script uses three methods to check the speed of connections to the file share prior to launching Access databases.
*  Response times to pings to any network domain name controller
*  Transfer a small file to and from a file share and the local machine several times
*  Transfer a single medium sized file (2mb) to and from the file share once.

# Why?
To date Microsoft Access databases can't be run directly on a fileshare over the internet. Connecting to MS Access databases over the internet does not work, and often corrupts the database. "It's important to understand that any time an Access client disconnects unexpectedly, it may set a "corruption flag" https://www.kzsoftware.com/articles/PreventAccessDatabaseCorruption

# CITRIX Connection
While connected to the file share over a VPN users must connect to a CITRIX virtual machine if they need to update a MS Access database on the network. When the script finds that a speed check fails the script launches a how to guide defined in the config.ini file i.e. [Use CITRIX to Run Access Database.pdf](/app/documentation/Use CITRIX to Run Access Database.pdf) and Internet Explorer to the CITRIX launcher.

# Generic Script
This script can be incorporated into any MS access database that has a backend on the network it is published to:

# Modify this script for your needs
Look for `[TODO]` comments in the [Config.ini](/app/Config.ini) and [Fileshare-Connection-Test.vbs](/app/Fileshare-Connection-Test.vbs) files

<<<<<<< HEAD
# Project
VBS, VBA, HTA
=======
# Problems and feedback 
Report any problems with this script as a [new issue](https://github.com/seakintruth/fileshare_connection_test_vbs/issues/new)
>>>>>>> 4f0f0c2b8c0fc9ad959522f4cf1ee3e8b6934c20
