These scripts were created for purpose and will need to be modified to work properly within whatever environment you put them in.

AServer.ps1 - Server script. Use NSSM or similar to create a service. The script is meant to run in a loop until killed.

AppStore.ps1 - Client script.

AppStore.ps1.cfg - Client configuration options

AppStore.ps1.ico - Client icon file

AppStore.vbs - Client wrapper script (if needed)

CreateAppStoreDB.ps1 - Database creation and initialization script

The scripts utilize the SQLite binary (downloaded separately) to access a shared database on a network drive.

The AServer.ps1 script should be run on the PDQ Deploy server as it needs to run PDQDeploy command-line options for various purposes
