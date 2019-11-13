InventorApiExamples

# Case 1: Vault Add-in login window messing with STA thread synchronization context
Copy contents of bin\debug to
C:\ProgramData\Autodesk\Inventor Addins\ReproduceVaultBug\

Alternatively change the *.addin file to point to the output directory of the build and copy only the *.addin file to the addins folder.


## Steps to reproduce:
1. Make sure auto-login is disabled in the Inventor Vault Add-in, if not, run Inventor, log out of Vault and then exit Inventor
2. Run Inventor and open the "Tools" ribbon tab.
3. Click all 4 buttons. All of them should open one or two small dialog windows.
4. Log in to Vault in Inventor
5. Click the "Fails After Login" button. This will throw an error. 
	If debugging, notice how the Synchronization Context is lost and the appartment state has changed to MTA after the awaited call
6. Click any of the other three buttons. All should succeed and open dialog windows.
	If debugging, observe the debug output to see how the Synchronization context and thread appartment state behaves.


# Case 2: Error when checking in *.dwg files in Inventor through custom code using Vault API
Copy the contents of bin\Debug to
C:\ProgramData\Autodesk\Inventor Addins\InvDwgCheckIn\

## How to reproduce:
1. Run Inventor, and make sure you're logged in to Vault through the Inventor add-in
2. Create a new drawing using the *.dwg template
3. Save the drawing somewhere in the Vault workspace
4. Under the "Drawing Check-In" ribbon tab, click the "Check In Drawing" button
Expected result:
An error should be thrown saying the file is used by another process

5. Create a new drawing, using the *.idw template this time
6. Save the drawing somewhere in the Vault workspace
7. Click the "Check In Drawing" button again
Expected result:
File should be checked in to Vault and a message box should say "Success!"
Refreshing the Vault browser in Inventor should reflect the successful check-in as well