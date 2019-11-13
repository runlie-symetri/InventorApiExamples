Copy contents of bin\debug to
C:\ProgramData\Autodesk\Inventor Addins\ReproduceVaultBug\

Alternatively change the *.addin file to point to the output directory of the build and copy only the *.addin file to the addins folder.


Steps to reproduce:
1. Make sure auto-login is disabled in the Inventor Vault Add-in, if not, run Inventor, log out of Vault and then exit Inventor
2. Run Inventor and open the "Tools" ribbon tab.
3. Click all 4 buttons. All of them should open one or two small dialog windows.
4. Log in to Vault in Inventor
5. Click the "Fails After Login" button. This will throw an error. 
	If debugging, notice how the Synchronization Context is lost and the appartment state has changed to MTA after the awaited call
6. Click any of the other three buttons. All should succeed and open dialog windows.
	If debugging, observe the debug output to see how the Synchronization context and thread appartment state behaves.