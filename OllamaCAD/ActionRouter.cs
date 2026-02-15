using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Newtonsoft.Json.Linq;
using SolidWorks.Interop.sldworks;

namespace OllamaCAD
{   
    /// <summary>
    /// Routes and executes JSON-based CAD actions inside SOLIDWORKS.
    /// 
    /// - Detects if a string contains a valid JSON object with an "actions" array.
    /// - Parses the JSON using Newtonsoft.Json.
    /// - Iterates through defined actions and executes corresponding SOLIDWORKS operations.
    /// - Currently supports placeholder actions like "create_box" and "fillet".
    /// - Logs unsupported actions and handles JSON parsing errors safely.
    /// 
    /// Acts as a bridge between LLM-generated structured output and SOLIDWORKS API commands.
    /// </summary>
    internal static class ActionRouter
    {
        public static bool LooksLikeJsonActions(string s)
        {
            if (string.IsNullOrWhiteSpace(s)) return false;
            string t = s.TrimStart();
            return t.StartsWith("{") && t.IndexOf("\"actions\"", StringComparison.OrdinalIgnoreCase) >= 0;
        }

        public static void TryExecute(ISldWorks app, string json, Action<string> log)
        {
            ModelDoc2 doc = (ModelDoc2)app.ActiveDoc;
            if (doc == null) { log("No active document.\r\n"); return; }

            try
            {
                JObject root = JObject.Parse(json);
                JToken actionsTok = root["actions"];
                if (actionsTok == null || actionsTok.Type != JTokenType.Array)
                {
                    log("No valid 'actions' array found.\r\n");
                    return;
                }

                foreach (JToken act in actionsTok)
                {
                    string name = (string)act["action"] ?? "";
                    JToken parameters = act["params"]; // optional

                    switch (name)
                    {
                        case "create_box":
                            log("Executing: create_box (TODO)\r\n");
                            break;

                        case "fillet":
                            log("Executing: fillet (TODO)\r\n");
                            break;

                        default:
                            log("Unsupported action: " + (string.IsNullOrEmpty(name) ? "(missing)" : name) + "\r\n");
                            break;
                    }
                }
            }
            catch (Exception ex)
            {
                log("JSON parse/exec error: " + ex.Message + "\r\n");
            }
        }
    }
}