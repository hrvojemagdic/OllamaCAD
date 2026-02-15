using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OllamaCAD
{   
    /// <summary>
    /// Builds structured prompts for the AI assistant by combining:
    /// - Current SOLIDWORKS document context (title, type, selection count)
    /// - The user's natural language request
    ///
    /// Also instructs the model to return strict JSON when model-changing
    /// actions are requested, enabling automated execution via ActionRouter.
    ///
    /// Acts as the bridge between CAD context and structured LLM output.
    /// </summary>
    internal static class SwPrompt
    {   
        /// <summary>
        /// Composes a prompt that embeds SOLIDWORKS context and enforces
        /// JSON-only output format for executable CAD actions.
        /// </summary>
        public static string Compose(string userText, SwContext ctx)
        {
            return
                "SOLIDWORKS context:\n" +
                "- Title: " + ctx.Title + "\n" +
                "- DocType: " + ctx.DocType + "\n" +
                "- SelectionCount: " + ctx.SelectionCount + "\n\n" +
                "User request:\n" + userText + "\n\n" +
                "If you want to change the model, return ONLY JSON:\n" +
                "{ \"actions\": [ { \"action\": \"create_box\", \"params\": { \"width_mm\": 50, \"height_mm\": 30, \"depth_mm\": 10 } } ] }\n";
        }
    }
}