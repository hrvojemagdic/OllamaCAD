using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;

namespace OllamaCAD
{   
    /// <summary>
    /// Lightweight snapshot of the current SOLIDWORKS document context.
    ///
    /// Contains:
    /// - Document title
    /// - Document type (part/assembly/drawing)
    /// - Current selection count
    ///
    /// Used to inject real-time CAD state into AI prompts.
    /// </summary>
    internal sealed class SwContext
    {
        public string Title;
        public int DocType;
        public int SelectionCount;
    }

    /// <summary>
    /// Builds a SwContext object from the active SOLIDWORKS session.
    ///
    /// Reads:
    /// - Active document
    /// - Document type
    /// - Current selection manager
    ///
    /// Provides minimal but useful context for AI-assisted responses.
    /// </summary>
    internal static class SwContextBuilder
    {   
        /// <summary>
        /// Creates a SwContext snapshot from the currently active SOLIDWORKS document.
        /// Returns a default placeholder context if no document is open.
        /// </summary>
        public static SwContext Build(ISldWorks app)
        {
            ModelDoc2 doc = (ModelDoc2)app.ActiveDoc;
            if (doc == null)
            {
                return new SwContext { Title = "(no document)", DocType = -1, SelectionCount = 0 };
            }

            SelectionMgr selMgr = (SelectionMgr)doc.SelectionManager;

            return new SwContext
            {
                Title = doc.GetTitle(),
                DocType = doc.GetType(),
                SelectionCount = selMgr != null ? selMgr.GetSelectedObjectCount2(-1) : 0
            };
        }
    }
}
