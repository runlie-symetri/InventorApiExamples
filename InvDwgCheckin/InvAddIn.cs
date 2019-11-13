using System;
using System.Drawing;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Windows;
using Autodesk.Connectivity.WebServices;
using Autodesk.DataManagement.Client.Framework.Currency;
using Autodesk.DataManagement.Client.Framework.Vault.Currency.Connections;
using Inventor;
using Application = Inventor.Application;
using Folder = Autodesk.DataManagement.Client.Framework.Vault.Currency.Entities.Folder;
using Path = System.IO.Path;

namespace InvDwgCheckIn
{
    [ComVisible(true)]
    [Guid(ADD_IN_GUID)]
    public class InvAddIn : ApplicationAddInServer
    {
        #region Class Fields
        public const string ADD_IN_GUID = "43D40537-ED29-45F9-A803-98B44CE802CE";
        public const string ADD_IN_CLS_GUID = "{" + ADD_IN_GUID + "}";

        private const string ButtonInternalName = "dwgCheckInTestButton";
        private const string ButtonIconResourceName = "InvDwgCheckIn.buttonIcon.ico";
        private const string RibbonTabInternalName = "dwgCheckInTestTab";
        private const string RibbonPanelInternalName = "dwgCheckInTestPanel";
        private const string DrawingRibbonId = "Drawing";

        private bool _ribbonInitialized;
        private Application _application;
        private ButtonDefinition _buttonDefinition;
        #endregion

        // This will fail for Inventor *.dwg files. For some reason the file handle is being kept open
        // by Inventor for *.dwg files, and not for *.idw files.
        // Steps to reproduce (Need to be logged in to Vault):
        // 1. Create a new drawing using the *.dwg template
        // 2. Save file within Vault workspace
        // 3. Hit the custom "Check in Drawing" command button in the ribbon
        // Expected result: File should be checked in
        // Actual result: IOException due to file being used by another process (Inventor.exe)
        //
        // Try the same procedure using the *.idw template, and the problem is no longer present
        private void ButtonDefinition_OnExecute(NameValueMap context)
        {
            string filePath = _application.ActiveDocument.FullFileName;
            if (string.IsNullOrWhiteSpace(filePath))
                return;

            Connection connection = Connectivity.Application.VaultBase.ConnectionManager.Instance.Connection;
            if (!connection?.IsConnected ?? false)
                return;

            // Get or create the relevant folder for the file in Vault
            Folder vaultFolder = GetOrCreateFolderFromLocalFilePath(filePath, connection);

            // Try to add the file
            try
            {
                connection?.FileManager.AddFile(
                    vaultFolder,
                    "Added by test add-in",
                    null,
                    null,
                    FileClassification.None,
                    false,
                    new FilePathAbsolute(filePath));

                MessageBox.Show($"Successfully checked in {Path.GetFileName(filePath)}!", "Success",
                    MessageBoxButton.OK, MessageBoxImage.Information);
            }
            catch (System.IO.IOException ex)
            {
                // This exception is thrown when trying to check in
                // Inventor *.dwg files. *.idw files work without issues
                MessageBox.Show(
                    $"IO Exception when checking in file for unknown reason. File should be available and writable at this point, but exception reporting that the file is being used by another process: {ex.Message}",
                    "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error: {ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void ApplicationEventsOnOnReady(EventTimingEnum beforeOrAfter, NameValueMap context, out HandlingCodeEnum handlingCode)
        {
            handlingCode = HandlingCodeEnum.kEventNotHandled;

            if (beforeOrAfter == EventTimingEnum.kAfter)
                AddCommand();

            handlingCode = HandlingCodeEnum.kEventHandled;
        }

        #region Interface Implementation

        public void Activate(ApplicationAddInSite addInSiteObject, bool firstTime)
        {
            _application = addInSiteObject.Application;

            if(firstTime)
                _application.ApplicationEvents.OnReady += ApplicationEventsOnOnReady;
        }

        public void Deactivate()
        {
            _application.ApplicationEvents.OnReady -= ApplicationEventsOnOnReady;
        }

        public void ExecuteCommand(int commandId)
        {
        }

        public dynamic Automation => null;

        #endregion

        #region Private Helpers

        private Folder GetOrCreateFolderFromLocalFilePath(string filePath, Connection connection)
        {
            FolderPathAbsolute workingFolder = connection.WorkingFoldersManager.GetWorkingFolder("$");
            string fileDirectory = Path.GetDirectoryName(filePath) ?? string.Empty;

            string subPath = string.Empty;
            for (int i = 0; i < fileDirectory.Length; i++)
            {
                if (workingFolder.FullPath.Length > i && workingFolder.FullPath[i] == fileDirectory[i])
                    continue;
                subPath += fileDirectory[i];
            }

            string vaultPath = $"$/{subPath.Replace('\\', '/')}";
            Autodesk.Connectivity.WebServices.Folder wsFolder = connection.WebServiceManager.DocumentService.FindFoldersByPaths(new[] { vaultPath }).First();
            return wsFolder.Id == -1
                ? GetOrCreateFolder(vaultPath, connection) : new Folder(connection, wsFolder);
        }

        private Folder GetOrCreateFolder(string path, Connection connection)
        {
            string[] pathParts = path.Split('/');
            Autodesk.Connectivity.WebServices.Folder lastAddedFolder = null;
            Autodesk.Connectivity.WebServices.Folder lastParent = connection.FolderManager.RootFolder;

            for (int i = 0; i < pathParts.Length; i++)
            {
                string folderName = pathParts[i];
                if (i > 0)
                    pathParts[i] = string.Join("/", pathParts[i - 1], pathParts[i]);
                Autodesk.Connectivity.WebServices.Folder[] folderResult = connection.WebServiceManager.DocumentService.FindFoldersByPaths(new[] { pathParts[i] });
                if (folderResult.First().Id == -1) // Folder doesn't exist in Vault. Let's add it
                {
                    try
                    {
                        lastAddedFolder =
                            connection.WebServiceManager.DocumentService.AddFolder(folderName, lastParent.Id, lastParent.IsLib);
                        lastParent = lastAddedFolder;
                    }
                    catch (Exception e)
                    {
                        MessageBox.Show($"Error: {e.Message}", "Error");
                    }

                }
                else
                    lastParent = folderResult.First();
            }

            return new Folder(connection, lastAddedFolder ?? lastParent);
        }

        private void AddCommand()
        {
            try
            {
                _buttonDefinition = _application.CommandManager.ControlDefinitions[ButtonInternalName] as ButtonDefinition;
            }
            catch (Exception)
            {
                _buttonDefinition =
                    _application.CommandManager.ControlDefinitions.AddButtonDefinition("Check In Drawing",
                        ButtonInternalName, CommandTypesEnum.kFileOperationsCmdType, Guid.NewGuid().ToString(), "Check in drawing to Vault", "Check in drawing to Vault", null, GetIconResource(ButtonIconResourceName));
            }

            if (!_ribbonInitialized &&
                _application.UserInterfaceManager.InterfaceStyle == InterfaceStyleEnum.kRibbonInterface)
            {
                Ribbon drawingRibbon = _application.UserInterfaceManager.Ribbons[DrawingRibbonId];
                RibbonTab ribbonTab =
                    drawingRibbon.RibbonTabs.Add("Drawing Check-In", RibbonTabInternalName, Guid.NewGuid().ToString());
                RibbonPanel ribbonPanel =
                    ribbonTab.RibbonPanels.Add("Check-in", RibbonPanelInternalName, Guid.NewGuid().ToString());
                ribbonPanel.CommandControls.AddButton(_buttonDefinition, true);

                _ribbonInitialized = true;
            }

            if (_buttonDefinition != null)
                _buttonDefinition.OnExecute += ButtonDefinition_OnExecute;
        }

        private object GetIconResource(string resourceName)
        {
            Assembly executingAssembly = Assembly.GetExecutingAssembly();
            Image icon;
            using (System.IO.Stream stream = executingAssembly.GetManifestResourceStream(resourceName))
            {
                if (stream == null)
                    return null;

                icon = Image.FromStream(stream);
            }

            return BitmapUtils.ToPictureDisp(icon);
        }

        #endregion
    }

    #region Bitmap Helper
    public sealed class BitmapUtils
    {
        public static object ToPictureDisp(Image bitmap)
        {
            return PictureConverter.BitmapToPictureDisp(bitmap);
        }
        

        private class PictureConverter : System.Windows.Forms.AxHost
        {
            private PictureConverter() : base("") { }

            internal static IPictureDisp BitmapToPictureDisp(Image bitmap)
            {
                return (IPictureDisp)GetIPictureDispFromPicture(bitmap);
            }
        }
    }
    #endregion
}