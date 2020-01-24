using System;
using System.Runtime.InteropServices;
using System.IO;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swpublished;

namespace _1225
{

    /// <summary>
    /// Our Solidworks taskpane add-in
    /// </summary>
    public class TaskPlane : ISwAddin
    {
        #region Private Members

        /// <summary>
        /// The cookie to the current instant of solidwork
        /// </summary>
        private int mSwCookie;

        /// <summary>
        /// The taskpane view for add-in
        /// </summary>
        private TaskpaneView mTaskpaneView;

        /// <summary>
        /// The UI control that is going to be inside the solidwork taskpane view
        /// </summary>
        private TaskPaneHostUI mTaskpaneHost;

        /// <summary>
        /// The current instance of solidworks application
        /// </summary>
        private SldWorks mSolidworksApplication;

        #endregion

        #region Public Members

        /// <summary>
        /// The unique Id to the taskpane used for registration in COM
        /// Since this is been called at the TaskPaneHostUI panel, this has to be Public
        /// </summary>

        public const string SwTaskPane_ProgId = "JustTaskPane";


        #endregion

        #region Solidworks Add-In Callbacks

        /// <summary>
        /// Called when solidworks has loaded our add in 
        /// </summary>
        /// <param name="ThisSW">The current Solidworks Instance</param>
        /// <param name="Cookie">The current Solidoworks cookie Id</param>
        /// <returns></returns>

        public bool ConnectToSW(object ThisSW, int Cookie)
        {
            // Store a reference to the current Solidworks instance
            mSolidworksApplication = (SldWorks)ThisSW;

            //store cookie Id
            mSwCookie = Cookie;

            // Setup callback info
            var ok = mSolidworksApplication.SetAddinCallbackInfo2(0, this, mSwCookie);
            LoadUI();
            return true; //Assume everything is okay..
        }


        /// <summary>
        /// Called when solidworks is about to unload your add in and disconnect logic 
        /// </summary>
        /// <returns></returns>

        public bool DisconnectFromSW()
        {
            UnloadUI();
            return true;
        }



        #endregion

        #region Create UI

        private void LoadUI()
        {
            var imagePath = Path.Combine(Path.GetDirectoryName(typeof(TaskPlane).Assembly.CodeBase).Replace(@"file:\", string.Empty), "Arrow_Small.png");
            // Create our Taskpane
            mTaskpaneView = mSolidworksApplication.CreateTaskpaneView2(imagePath, "Woo! My first SwAddin");

            // Load our UI into the taskpane
            mTaskpaneHost = (TaskPaneHostUI)mTaskpaneView.AddControl(TaskPlane.SwTaskPane_ProgId, string.Empty);
        }

        private void UnloadUI()
        {
            mTaskpaneHost = null;

            // Remove taskpane view
            mTaskpaneView.DeleteView();

            // Release COM reference and cleanup memory
            Marshal.ReleaseComObject(mTaskpaneView);

            mTaskpaneView = null;
        }
        #endregion

        #region COM Registration

        [ComRegisterFunction()]
        private static void ComRegister(Type t)
        {
            var keyPath = string.Format(@"SoftWare\SolidWorks\AddIns\{0:b}", t.GUID); //GuID for this application

            using (var rk = Microsoft.Win32.Registry.LocalMachine.CreateSubKey(keyPath))
            {
                // Load add-in when Solidworks Opens
                rk.SetValue(null, 1);

                rk.SetValue("Title", "My SwAddIn");
                rk.SetValue("Description", "All your pixels are belong to us!");
            }
        }

        [ComUnregisterFunction()]
        private static void Comregister(Type t)
        {
            var keyPath = string.Format(@"SoftWare\SolidWorks\AddIns\{0:b}", t.GUID); //GuID for this application

            Microsoft.Win32.Registry.LocalMachine.DeleteSubKeyTree(keyPath);
        }




        #endregion
    }

}

