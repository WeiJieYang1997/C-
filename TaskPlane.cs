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
        /// </summary>

        public const string SwTaskPane_ProgId = "JustTaskPane";
        #endregion
    }
}
