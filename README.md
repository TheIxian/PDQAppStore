These scripts were created for purpose and will need to be modified to work properly within whatever environment you put them in.

AServer.ps1 - Server script. Use NSSM or similar to create a service. The script is meant to run in a loop until killed.

AppStore.ps1 - Client script.

AppStore.ps1.cfg - Client configuration options

AppStore.ps1.ico - Client icon file

AppStore.vbs - Client wrapper script (if needed)

CreateAppStoreDB.ps1 - Database creation and initialization script

The scripts utilize the SQLite binary (downloaded separately) to access a shared database on a network drive.

The AServer.ps1 script should be run on the PDQ Deploy server as it needs to run PDQDeploy command-line options for various purposes

AServer.ps1 looks for a particular format in the package name and will only add items to the packages database which conform to this format:

[AppStore] (Category) Name of package

The AppStore.ps1 script is meant to be run from a read-only network share

When AppStore.ps1 is executed by a member of a group specified in the $adminControlsGroup variable, an extra button is added to the GUI which allows for various admin-level functions like editing the notes field, scheduling an import of packages from DPQ Deploy, editing AppStore packages, etc.
