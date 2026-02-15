using CodeStack.SwEx.AddIn;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using SolidWorks.Interop.sldworks;
using CodeStack.SwEx.AddIn.Attributes;
using System.Windows.Forms;
using SolidWorks.Interop.swconst;
using SolidWorks.Interop.swpublished;

namespace OllamaCAD
{   
    /// <summary>
    /// Main entry point for the OllamaCAD SOLIDWORKS add-in.
    /// 
    /// - Registers the add-in as a COM-visible component.
    /// - Connects to the active SOLIDWORKS instance.
    /// - Sets callback information for SOLIDWORKS.
    /// - Creates and displays the custom Task Pane UI.
    /// - Properly disposes resources when the add-in is unloaded.
    /// 
    /// Acts as the integration layer between SOLIDWORKS and the AI-powered Task Pane interface.
    /// </summary>
    [ComVisible(true)]
    [Guid("D21CDAF8-30C5-46DE-9B44-2386572E1D43")]
    public class AddIn : ISwAddin
    {
        private ISldWorks _swApp;
        private int _cookie;
        private TaskpaneHost _taskpaneHost;
        
        /// <summary>
        /// Called when the add-in is loaded into SOLIDWORKS.
        /// Initializes connection and shows the Task Pane.
        /// </summary>
        public bool ConnectToSW(object ThisSW, int Cookie)
        {
            _cookie = Cookie;

            try
            {
                // IMPORTANT: bind to the actual running SOLIDWORKS instance
                _swApp = (ISldWorks)Marshal.GetActiveObject("SldWorks.Application");

                _swApp.SetAddinCallbackInfo2(0, this, _cookie);

                _taskpaneHost = new TaskpaneHost(_swApp);
                _taskpaneHost.ShowTaskpane();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString(), "ConnectToSW error");
            }

            return true;
        }
        
        /// <summary>
        /// Called when the add-in is unloaded from SOLIDWORKS.
        /// Cleans up Task Pane and releases references.
        /// </summary>
        public bool DisconnectFromSW()
        {
            try { _taskpaneHost?.Dispose(); } catch { }
            _taskpaneHost = null;
            _swApp = null;
            return true;
        }
    }
}