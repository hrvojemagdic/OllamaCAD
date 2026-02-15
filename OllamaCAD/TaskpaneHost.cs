using System;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Windows.Forms;
using SolidWorks.Interop.sldworks;


namespace OllamaCAD
{   
    /// <summary>
    /// Creates, hosts, and cleans up the SOLIDWORKS Task Pane for the OllamaCAD add-in.
    ///
    /// Responsibilities:
    /// - Creates a TaskpaneView using COM late-binding to support multiple SOLIDWORKS versions
    ///   (tries CreateTaskpaneView2, then CreateTaskpaneView3).
    /// - Loads the task pane icon from the add-in DLL directory (OllamaCAD.png).
    /// - Instantiates the ChatPaneControl UI and embeds it into the Task Pane via DisplayWindowFromHandle.
    /// - Provides helper methods for debugging task pane overloads and display methods.
    /// - Disposes the TaskpaneView (DeleteView) and UI when the add-in is unloaded.
    ///
    /// Uses object + reflection intentionally to avoid interop casting/signature differences between SW versions.
    /// </summary>
    internal sealed class TaskpaneHost : IDisposable
    {

        private readonly ISldWorks _swApp;

        // Keep as object to avoid COM cast issues
        private object _taskpaneObj;
        private UserControl _ui;

        public TaskpaneHost(ISldWorks swApp)
        {
            _swApp = swApp;
        }

        /// <summary>
        /// Creates the SOLIDWORKS Task Pane and embeds the ChatPaneControl into it.
        /// </summary>
        public void ShowTaskpane()
        {
            string dllDir = Path.GetDirectoryName(typeof(TaskpaneHost).Assembly.Location);
            string iconPath = Path.Combine(dllDir, "OllamaCAD.png");
            if (!File.Exists(iconPath)) iconPath = "";

            // Create via COM late-binding to support different SW versions
            object tpObj = null;

            try
            {
                // Try CreateTaskpaneView2(string iconPath, string caption)
                tpObj = _swApp.GetType().InvokeMember(
                    "CreateTaskpaneView2",
                    BindingFlags.InvokeMethod,
                    null,
                    _swApp,
                    new object[] { iconPath, "Ollama Assistant" }
                );
            }
            catch
            {
                // Try CreateTaskpaneView3(string iconPath, string caption, object handler)
                try
                {
                    tpObj = _swApp.GetType().InvokeMember(
                        "CreateTaskpaneView3",
                        BindingFlags.InvokeMethod,
                        null,
                        _swApp,
                        new object[] { iconPath, "Ollama Assistant", null }
                    );
                }
                catch
                {
                    tpObj = null;
                }
            }

            if (tpObj == null)
                throw new Exception("Could not create TaskpaneView via CreateTaskpaneView2/3. (tpObj is null)");

            _ui = new ChatPaneControl(_swApp);
            _ui.CreateControl();

            // DisplayWindowFromHandle on the returned TaskpaneView object
            tpObj.GetType().InvokeMember(
                "DisplayWindowFromHandle",
                BindingFlags.InvokeMethod,
                null,
                tpObj,
                new object[] { _ui.Handle.ToInt32() }
            );

            _taskpaneObj = tpObj;
        }

        private void ShowCreateTaskpaneOverloads()
        {
            try
            {
                string text = string.Join("\r\n",
                    _swApp.GetType()
                        .GetMethods()
                        .Where(m => m.Name.IndexOf("CreateTaskpaneView", StringComparison.OrdinalIgnoreCase) >= 0)
                        .Select(m => m.ToString())
                        .Distinct()
                        .ToArray()
                );

                MessageBox.Show(
                    string.IsNullOrWhiteSpace(text) ? "(no CreateTaskpaneView* methods found on this COM object)" : text,
                    "CreateTaskpaneView overloads"
                );
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Overload listing failed");
            }
        }

        private object TryCreateTaskpaneObjectWithIconPath()
        {
            string dllDir = Path.GetDirectoryName(typeof(TaskpaneHost).Assembly.Location);
            string iconPath = Path.Combine(dllDir, "OllamaCAD.png"); // your icon filename

            if (!File.Exists(iconPath))
            {
                // still attempt with empty icon path (some versions accept it)
                iconPath = "";
            }

            // Find any CreateTaskpaneView* overload that takes (string, string) or (string, string, object)
            var methods = _swApp.GetType().GetMethods()
                .Where(m => m.Name.StartsWith("CreateTaskpaneView", StringComparison.OrdinalIgnoreCase))
                .ToArray();

            foreach (var m in methods)
            {
                var ps = m.GetParameters();

                // (string iconPath, string caption)
                if (ps.Length == 2 &&
                    ps[0].ParameterType == typeof(string) &&
                    ps[1].ParameterType == typeof(string))
                {
                    try
                    {
                        object res = m.Invoke(_swApp, new object[] { iconPath, "Ollama Assistant" });
                        if (res != null) return res;
                    }
                    catch { }
                }

                // (string iconPath, string caption, object handler)
                if (ps.Length == 3 &&
                    ps[0].ParameterType == typeof(string) &&
                    ps[1].ParameterType == typeof(string))
                {
                    try
                    {
                        object res = m.Invoke(_swApp, new object[] { iconPath, "Ollama Assistant", null });
                        if (res != null) return res;
                    }
                    catch { }
                }
            }

            return null;
        }

        private void DisplayInTaskpaneObject(object taskpaneObj, IntPtr hwnd)
        {
            if (taskpaneObj == null)
                throw new ArgumentNullException("taskpaneObj");

            Type t = taskpaneObj.GetType();

            // Most common SW2020 method
            MethodInfo m = t.GetMethod("DisplayWindowFromHandle");
            if (m != null)
            {
                m.Invoke(taskpaneObj, new object[] { hwnd.ToInt32() });
                return;
            }

            // Some interops expose x64 variant
            MethodInfo m64 = t.GetMethod("DisplayWindowFromHandlex64");
            if (m64 != null)
            {
                m64.Invoke(taskpaneObj, new object[] { hwnd.ToInt64() });
                return;
            }

            throw new NotSupportedException("Returned taskpane object has no DisplayWindowFromHandle/Handlex64 method.");
        }

        /// <summary>
        /// Deletes the TaskpaneView and disposes the hosted UI control.
        /// </summary>
        public void Dispose()
        {
            try
            {
                if (_taskpaneObj != null)
                {
                    _taskpaneObj.GetType().InvokeMember(
                        "DeleteView",
                        BindingFlags.InvokeMethod,
                        null,
                        _taskpaneObj,
                        new object[0]
                    );
                }
            }
            catch { }

            _taskpaneObj = null;

            try { _ui?.Dispose(); } catch { }
            _ui = null;
        }
    }
}
