Copy the contents of bin\Debug to
C:\ProgramData\Autodesk\Inventor Addins\InvDwgCheckIn\

How to reproduce:
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