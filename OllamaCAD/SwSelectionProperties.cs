using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;

namespace OllamaCAD
{
    /// <summary>
    /// Builds a structured, text-only property block for currently selected
    /// SOLIDWORKS components or parts.
    ///
    /// Purpose:
    /// - Extracts material, mass properties, bounding box, units, and selected
    ///   custom properties from selected components.
    /// - Handles assembly selections (Component2) and part fallback logic.
    /// - Uses interop-safe patterns (dynamic where needed) for compatibility
    ///   across SOLIDWORKS versions.
    /// - Formats output as "Label=value" pairs for reliable LLM prompt injection.
    ///
    /// Designed to be appended to the system prompt so the AI can reason about
    /// selected geometry and metadata without direct API access.
    /// </summary>
    internal static class SwSelectionProperties
    {   
        /// <summary>
        /// Returns a formatted context block describing the currently selected
        /// components in the active document. Returns empty string if nothing is selected.
        /// </summary>
        internal static string BuildSelectedPartsContext(ISldWorks swApp)
        {
            if (swApp == null) return "";

            ModelDoc2 doc = null;
            ISelectionMgr selMgr = null;

            try
            {
                doc = swApp.ActiveDoc as ModelDoc2;
                if (doc == null) return "";
                selMgr = doc.SelectionManager as ISelectionMgr;
                if (selMgr == null) return "";

                int count = 0;
                try { count = selMgr.GetSelectedObjectCount2(-1); } catch { count = 0; }
                if (count <= 0) return "";

                var rows = new List<Row>();

                for (int i = 1; i <= count; i++)
                {
                    int selType = 0;
                    try { selType = selMgr.GetSelectedObjectType3(i, -1); } catch { selType = 0; }

                    // Prefer component selections (assembly context)
                    // NOTE: SOLIDWORKS enum is swSelCOMPONENTS (plural). Some code samples mention swSelCOMPONENT,
                    // but that member is not present in many interop versions (including SW 2020).
                    if (selType == (int)swSelectType_e.swSelCOMPONENTS)
                    {
                        object obj = null;
                        try { obj = selMgr.GetSelectedObject6(i, -1); } catch { obj = null; }
                        var comp = obj as Component2;
                        if (comp == null) continue;

                        Row r = CollectFromComponent(swApp, comp);
                        if (r != null) rows.Add(r);
                        continue;
                    }

                    // If user selected a part entity while inside a Part document, include active doc as a fallback
                    // but only once.
                    if (doc.GetType() == (int)swDocumentTypes_e.swDocPART)
                    {
                        if (rows.Count == 0)
                        {
                            Row r = CollectFromModelDoc(doc, componentPath: doc.GetTitle() ?? "(active part)",
                                filePathOverride: doc.GetPathName() ?? "", suppressed: false, configOverride: null);
                            if (r != null) rows.Add(r);
                        }
                    }
                }

                if (rows.Count == 0) return "";

                return Format(rows);
            }
            catch
            {
                return "";
            }
            finally
            {
                // do NOT release doc/selMgr - owned by SOLIDWORKS
            }
        }

        private sealed class Row
        {
            public string ComponentPath = "";
            public string FilePath = "";
            public string DocType = "";
            public string Config = "";
            public bool Suppressed;
            public string Units = "";
            public string MaterialDb = "";
            public string MaterialName = "";
            public double Density = double.NaN;
            public double Mass = double.NaN;
            public double CoG_X = double.NaN;
            public double CoG_Y = double.NaN;
            public double CoG_Z = double.NaN;
            public double Ixx = double.NaN;
            public double Iyy = double.NaN;
            public double Izz = double.NaN;
            public double Volume = double.NaN;
            public double SurfaceArea = double.NaN;
            public double BBX_X = double.NaN;
            public double BBX_Y = double.NaN;
            public double BBX_Z = double.NaN;
            public string PROP_PartNo = "";
            public string PROP_Description = "";
            public string PROP_Revision = "";
            public string PROP_Finish = "";
        }

        private static Row CollectFromComponent(ISldWorks swApp, Component2 comp)
        {
            if (comp == null) return null;

            bool suppressed = false;
            try { suppressed = comp.IsSuppressed(); } catch { suppressed = false; }

            string cfg = "";
            try { cfg = comp.ReferencedConfiguration ?? ""; } catch { cfg = ""; }

            string compPath = BuildComponentPath(comp);
            string filePath = "";
            try { filePath = comp.GetPathName() ?? ""; } catch { filePath = ""; }

            // Lightweight components may not return ModelDoc2.
            ModelDoc2 cDoc = null;
            try { cDoc = comp.GetModelDoc2() as ModelDoc2; } catch { cDoc = null; }

            if (cDoc == null)
            {
                // If it is already open, get by name
                if (!string.IsNullOrWhiteSpace(filePath))
                {
                    try { cDoc = swApp.GetOpenDocumentByName(filePath) as ModelDoc2; } catch { cDoc = null; }
                }
            }

            if (cDoc == null)
            {
                // Unresolved selection
                return new Row
                {
                    ComponentPath = compPath,
                    FilePath = filePath,
                    DocType = "Unresolved",
                    Config = cfg,
                    Suppressed = suppressed,
                    Units = ""
                };
            }

            // Collect. Do NOT release cDoc (owned by SW when returned by GetModelDoc2 / GetOpenDocumentByName).
            return CollectFromModelDoc(cDoc, compPath, filePath, suppressed, cfg);
        }

        private static Row CollectFromModelDoc(ModelDoc2 model, string componentPath, string filePathOverride, bool suppressed, string configOverride)
        {
            if (model == null) return null;

            var row = new Row();
            row.ComponentPath = componentPath ?? "";
            row.FilePath = !string.IsNullOrWhiteSpace(filePathOverride) ? filePathOverride : (model.GetPathName() ?? "");
            row.DocType = GetDocTypeName(model);
            row.Config = !string.IsNullOrWhiteSpace(configOverride)
                ? configOverride
                : (model.ConfigurationManager != null && model.ConfigurationManager.ActiveConfiguration != null
                    ? (model.ConfigurationManager.ActiveConfiguration.Name ?? "")
                    : "");
            row.Suppressed = suppressed;
            row.Units = TryGetUnitsString(model);

            // Material
            string db, mat;
            TryGetMaterialCompat(model, row.Config, out db, out mat);
            row.MaterialDb = db ?? "";
            row.MaterialName = mat ?? "";

            // Mass props
            TryMassPropsCompat(model, row);

            // Bounding box
            TryBoundingBoxPart(model, row);

            // Custom props
            TryReadExistingPropsCompat(model, row.Config, row);

            return row;
        }

        private static string Format(List<Row> rows)
        {
            var sb = new StringBuilder();
            sb.AppendLine("SelectedComponents: " + rows.Count.ToString(CultureInfo.InvariantCulture));
            sb.AppendLine("For each selected component, properties are listed as Label=value.");
            sb.AppendLine();

            for (int i = 0; i < rows.Count; i++)
            {
                var r = rows[i];
                sb.AppendLine("-- Selected[" + i.ToString(CultureInfo.InvariantCulture) + "] --");

                sb.AppendLine("ComponentPath=" + Safe(r.ComponentPath));
                sb.AppendLine("FilePath=" + Safe(r.FilePath));
                sb.AppendLine("DocType=" + Safe(r.DocType));
                sb.AppendLine("Config=" + Safe(r.Config));
                sb.AppendLine("Suppressed=" + (r.Suppressed ? "True" : "False"));
                sb.AppendLine("Units=" + Safe(r.Units));
                sb.AppendLine("MaterialDb=" + Safe(r.MaterialDb));
                sb.AppendLine("MaterialName=" + Safe(r.MaterialName));
                sb.AppendLine("Density=" + F(r.Density));
                sb.AppendLine("Mass=" + F(r.Mass));
                sb.AppendLine("CoG_X=" + F(r.CoG_X));
                sb.AppendLine("CoG_Y=" + F(r.CoG_Y));
                sb.AppendLine("CoG_Z=" + F(r.CoG_Z));
                sb.AppendLine("Ixx=" + F(r.Ixx));
                sb.AppendLine("Iyy=" + F(r.Iyy));
                sb.AppendLine("Izz=" + F(r.Izz));
                sb.AppendLine("Volume=" + F(r.Volume));
                sb.AppendLine("SurfaceArea=" + F(r.SurfaceArea));
                sb.AppendLine("BBX_X=" + F(r.BBX_X));
                sb.AppendLine("BBX_Y=" + F(r.BBX_Y));
                sb.AppendLine("BBX_Z=" + F(r.BBX_Z));
                sb.AppendLine("PROP_PartNo=" + Safe(r.PROP_PartNo));
                sb.AppendLine("PROP_Description=" + Safe(r.PROP_Description));
                sb.AppendLine("PROP_Revision=" + Safe(r.PROP_Revision));
                sb.AppendLine("PROP_Finish=" + Safe(r.PROP_Finish));
                sb.AppendLine();
            }

            return sb.ToString().TrimEnd();
        }

        private static string Safe(string s) => s ?? "";

        private static string F(double v)
        {
            if (double.IsNaN(v) || double.IsInfinity(v)) return "";
            return v.ToString("G", CultureInfo.InvariantCulture);
        }

        private static string BuildComponentPath(Component2 c)
        {
            try
            {
                var parts = new List<string>();
                Component2 cur = c;
                while (cur != null)
                {
                    parts.Add(cur.Name2 ?? "(?)");
                    cur = cur.GetParent() as Component2;
                }
                parts.Reverse();
                return string.Join("^", parts);
            }
            catch
            {
                return c != null ? (c.Name2 ?? "(component)") : "(component)";
            }
        }

        private static string TryGetUnitsString(ModelDoc2 doc)
        {
            try
            {
                int unit = doc.GetUserPreferenceIntegerValue((int)swUserPreferenceIntegerValue_e.swUnitSystem);
                switch (unit)
                {
                    case (int)swUnitSystem_e.swUnitSystem_MMGS: return "MMGS";
                    case (int)swUnitSystem_e.swUnitSystem_CGS: return "CGS";
                    case (int)swUnitSystem_e.swUnitSystem_MKS: return "MKS";
                    case (int)swUnitSystem_e.swUnitSystem_IPS: return "IPS";
                    default: return "Unknown";
                }
            }
            catch { return ""; }
        }

        private static string GetDocTypeName(ModelDoc2 doc)
        {
            int t = 0;
            try { t = doc.GetType(); } catch { t = 0; }

            if (t == (int)swDocumentTypes_e.swDocPART) return "Part";
            if (t == (int)swDocumentTypes_e.swDocASSEMBLY) return "Assembly";
            if (t == (int)swDocumentTypes_e.swDocDRAWING) return "Drawing";
            return "Unknown";
        }

        private static void TryGetMaterialCompat(ModelDoc2 doc, string configName, out string db, out string matName)
        {
            db = "";
            matName = "";

            try
            {
                dynamic ext = doc.Extension;

                try { ext.GetMaterialPropertyName2(configName, out db, out matName); return; } catch { }
                try { ext.GetMaterialPropertyName(configName, out db, out matName); return; } catch { }
            }
            catch
            {
                // leave blank
            }
        }

        private static void TryMassPropsCompat(ModelDoc2 doc, Row row)
        {
            try
            {
                IModelDocExtension ext = doc.Extension;
                MassProperty mp = ext.CreateMassProperty() as MassProperty;
                if (mp == null) return;

                row.Mass = mp.Mass;

                double[] cog = mp.CenterOfMass as double[];
                if (cog != null && cog.Length >= 3)
                {
                    row.CoG_X = cog[0];
                    row.CoG_Y = cog[1];
                    row.CoG_Z = cog[2];
                }

                row.Volume = mp.Volume;
                row.SurfaceArea = mp.SurfaceArea;

                if (row.Volume > 0 && !double.IsNaN(row.Mass))
                    row.Density = row.Mass / row.Volume;

                try
                {
                    dynamic d = mp;
                    double[] p = d.PrincipalMomentsOfInertia;
                    if (p != null && p.Length >= 3)
                    {
                        row.Ixx = p[0];
                        row.Iyy = p[1];
                        row.Izz = p[2];
                    }
                }
                catch { }
            }
            catch { }
        }

        private static void TryBoundingBoxPart(ModelDoc2 doc, Row row)
        {
            try
            {
                if (doc.GetType() != (int)swDocumentTypes_e.swDocPART) return;
                PartDoc part = doc as PartDoc;
                if (part == null) return;

                double[] box = part.GetPartBox(true) as double[];
                if (box != null && box.Length >= 6)
                {
                    row.BBX_X = box[3] - box[0];
                    row.BBX_Y = box[4] - box[1];
                    row.BBX_Z = box[5] - box[2];
                }
            }
            catch { }
        }


        private static void TryReadExistingPropsCompat(ModelDoc2 doc, string cfg, Row row)
        {
            try
            {
                CustomPropertyManager cpm = doc.Extension.CustomPropertyManager[cfg ?? ""];

                row.PROP_PartNo = GetPropCompat(cpm, "PartNo");
                row.PROP_Description = GetPropCompat(cpm, "Description");
                row.PROP_Revision = GetPropCompat(cpm, "Revision");
                row.PROP_Finish = GetPropCompat(cpm, "Finish");
            }
            catch { }
        }

        private static string GetPropCompat(CustomPropertyManager cpm, string name)
        {
            try
            {
                string valOut, resolved;
                cpm.Get4(name, false, out valOut, out resolved);
                return !string.IsNullOrWhiteSpace(resolved) ? resolved : (valOut ?? "");
            }
            catch
            {
                return "";
            }
        }
    }
}
