using System;
using System.Diagnostics;
using System.Drawing;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Threading;
using Inventor;
using Application = Inventor.Application;

namespace ReproduceVaultBug
{
    [ComVisible(true)]
    [Guid("A8449AF5-FFB8-48CA-A26D-73F0D63C2F20")]
    public class MyAddinServer : ApplicationAddInServer
    {
        private ButtonDefinition _failsButton, _worksButton1, _worksButton2, _worksButton3;
        private Application _application;
        
        // Something happens to the UI thread when the Vault add-in is installed and you open the login window 
        public void Activate(ApplicationAddInSite AddInSiteObject, bool FirstTime)
        {
            _application = AddInSiteObject.Application;
            if(FirstTime)
                AddCommands();
        }

        //This fails after Vault login
        private async void MyButtonOnOnExecute(NameValueMap context)
        {
            try
            {
                Debug.WriteLine($"Synchronization Context: {SynchronizationContext.Current}");
                Debug.WriteLine($"Current Thread Id: {Thread.CurrentThread.ManagedThreadId} ApartmentState: {Thread.CurrentThread.GetApartmentState()}");

                await Task.Delay(1000).ConfigureAwait(true);

                Debug.WriteLine($"Synchronization Context: {SynchronizationContext.Current}");
                Debug.WriteLine($"Current Thread Id: {Thread.CurrentThread.ManagedThreadId} ApartmentState: {Thread.CurrentThread.GetApartmentState()}");

                new MyDialog().ShowDialog();
            }
            catch (Exception ex)
            {
               MessageBox.Show($"Exception caught: {ex.Message}", "Error", MessageBoxButton.OK,
                    MessageBoxImage.Error);
            }
        }

        // This works after Vault login
        private void MyButtonOnOnExecute2(NameValueMap context)
        {

            var dialog = new MyDialog();
            dialog.Loaded += async (sender, args) =>
            {
                Debug.WriteLine($"Synchronization Context: {SynchronizationContext.Current}");
                Debug.WriteLine($"Current Thread Id: {Thread.CurrentThread.ManagedThreadId} ApartmentState: {Thread.CurrentThread.GetApartmentState()}");

                await Task.Delay(1000).ConfigureAwait(true);

                Debug.WriteLine($"Synchronization Context: {SynchronizationContext.Current}");
                Debug.WriteLine($"Current Thread Id: {Thread.CurrentThread.ManagedThreadId} ApartmentState: {Thread.CurrentThread.GetApartmentState()}");

                new MyDialog().ShowDialog();
            };
            dialog.ShowDialog();
        }

        // This works after Vault login
        private async void MyButtonOnOnExecute3(NameValueMap context)
        {
            Dispatcher dispatcher = Dispatcher.CurrentDispatcher;

            await dispatcher.InvokeAsync(async () =>
            {
                Debug.WriteLine($"Synchronization Context: {SynchronizationContext.Current}");
                Debug.WriteLine(
                    $"Current Thread Id: {Thread.CurrentThread.ManagedThreadId} ApartmentState: {Thread.CurrentThread.GetApartmentState()}");

                await Task.Delay(1000).ConfigureAwait(true);

                Debug.WriteLine($"Synchronization Context: {SynchronizationContext.Current}");
                Debug.WriteLine(
                    $"Current Thread Id: {Thread.CurrentThread.ManagedThreadId} ApartmentState: {Thread.CurrentThread.GetApartmentState()}");

                new MyDialog().ShowDialog();
            });
        }

        // This works after Vault login
        private async void MyButtonOnOnExecute4(NameValueMap context)
        {
            var dispatcher = Dispatcher.CurrentDispatcher;

            Debug.WriteLine($"Synchronization Context: {SynchronizationContext.Current}");
            Debug.WriteLine(
                $"Current Thread Id: {Thread.CurrentThread.ManagedThreadId} ApartmentState: {Thread.CurrentThread.GetApartmentState()}");

            await Task.Delay(1000).ConfigureAwait(true);

            Debug.WriteLine($"Synchronization Context: {SynchronizationContext.Current}");
            Debug.WriteLine(
                $"Current Thread Id: {Thread.CurrentThread.ManagedThreadId} ApartmentState: {Thread.CurrentThread.GetApartmentState()}");

            await dispatcher.InvokeAsync(async () =>
            {
                Debug.WriteLine($"Synchronization Context: {SynchronizationContext.Current}");
                Debug.WriteLine(
                    $"Current Thread Id: {Thread.CurrentThread.ManagedThreadId} ApartmentState: {Thread.CurrentThread.GetApartmentState()}");

                new MyDialog().ShowDialog();
            });
        }

        public void Deactivate()
        {
            _failsButton.OnExecute -= MyButtonOnOnExecute;
            _worksButton1.OnExecute -= MyButtonOnOnExecute2;
            _worksButton2.OnExecute -= MyButtonOnOnExecute3;
            _worksButton3.OnExecute -= MyButtonOnOnExecute4;
        }

        public void ExecuteCommand(int CommandID) { }

        public object Automation { get; }

        //This property uses reflection to get the value for the GuidAttribute attached to the class.
        public string Guid
        {
            get
            {
                object[] customAttributes = GetType().GetCustomAttributes(typeof(GuidAttribute), false);
                var guidAttribute = (GuidAttribute)customAttributes[0];
                string guid = "{" + guidAttribute.Value + "}";

                return guid;
            }
        }

        private void AddCommands()
        {
            // Create the button definitions
            _failsButton = _application.CommandManager.ControlDefinitions.AddButtonDefinition(
                "Fails After Login",
                "failsAfterLoginButton", 
                CommandTypesEnum.kQueryOnlyCmdType,
                Guid,
                "Should fail if clicked after manual Vault login",
                "Should fail if clicked after manual Vault login",
                null,
                GetIconResource("ReproduceVaultBug.negative.ico"));
            _failsButton.OnExecute += MyButtonOnOnExecute;

            _worksButton1 = _application.CommandManager.ControlDefinitions.AddButtonDefinition(
                "Works After Login",
                "worksAfterLoginButton1",
                CommandTypesEnum.kQueryOnlyCmdType,
                Guid,
                "Should work if clicked after manual Vault login",
                "Should work if clicked after manual Vault login",
                null,
                GetIconResource("ReproduceVaultBug.positive.ico"));
            _worksButton1.OnExecute += MyButtonOnOnExecute2;

            _worksButton2 = _application.CommandManager.ControlDefinitions.AddButtonDefinition(
                "Works After Login",
                "worksAfterLoginButton2",
                CommandTypesEnum.kQueryOnlyCmdType,
                Guid,
                "Should work if clicked after manual Vault login",
                "Should work if clicked after manual Vault login",
                null,
                GetIconResource("ReproduceVaultBug.positive.ico"));
            _worksButton2.OnExecute += MyButtonOnOnExecute3;

            _worksButton3 = _application.CommandManager.ControlDefinitions.AddButtonDefinition(
                "Works After Login",
                "worksAfterLoginButton3",
                CommandTypesEnum.kQueryOnlyCmdType,
                Guid,
                "Should work if clicked after manual Vault login",
                "Should work if clicked after manual Vault login",
                null,
                GetIconResource("ReproduceVaultBug.positive.ico"));
            _worksButton3.OnExecute += MyButtonOnOnExecute4;

            RibbonTab tab = _application.UserInterfaceManager.Ribbons["ZeroDoc"].RibbonTabs["id_TabTools"];
            RibbonPanel panel;
            try
            {
                panel = tab.RibbonPanels["id_ReproduceVaultBug"];
            }
            catch (Exception)
            {
                panel = tab.RibbonPanels.Add("ReproduceVaultBug", "id_ReproduceVaultBug", Guid);
            }

            panel.CommandControls.AddButton(_failsButton, true);
            panel.CommandControls.AddButton(_worksButton1, true);
            panel.CommandControls.AddButton(_worksButton2, true);
            panel.CommandControls.AddButton(_worksButton3, true);
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
