using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using SwSheet = SolidWorks.Interop.sldworks.Sheet;
using OxSheet = DocumentFormat.OpenXml.Spreadsheet.Sheet;


namespace OllamaCAD
{   
    /// <summary>
    /// Exports SOLIDWORKS model/assembly metadata to an Excel (.xlsx) report and can import edited
    /// custom properties back into SOLIDWORKS.
    ///
    /// Export:
    /// - Collects key document info (type, title, path, active configuration, units, timestamp).
    /// - For assemblies, iterates components and captures properties such as:
    ///   file path, suppression state, material, mass properties (mass/CoG/inertia), volume/area,
    ///   bounding box, feature counts, and selected editable custom properties (PROP_* columns).
    /// - Writes a standalone .xlsx using OpenXML (no Microsoft Excel required).
    ///
    /// Import:
    /// - Reads the "Components" sheet from the exported .xlsx.
    /// - Requires "FilePath" and "Config" columns to locate target documents/configurations.
    /// - Applies any columns prefixed with "PROP_" as custom properties (adds or replaces values).
    ///
    /// Notes:
    /// - Uses interop-safe patterns (dynamic where needed) to handle SOLIDWORKS API signature differences.
    /// - Intended for batch editing of part/assembly custom properties via Excel.
    /// </summary>
    internal static class SwAssemblyExcelReport
    {
        // -------------------------
        // Public API
        // -------------------------

        /// <summary>Exports the active SOLIDWORKS document to an OpenXML .xlsx report and returns the saved file path.</summary>
        public static string ExportActiveDocToExcel(ISldWorks swApp, string outputFolder, Action<string> log)
{
    ModelDoc2 doc = swApp != null ? (swApp.ActiveDoc as ModelDoc2) : null;
    if (doc == null) throw new InvalidOperationException("No active document.");

    Directory.CreateDirectory(outputFolder);

    string stamp = DateTime.Now.ToString("yyyyMMdd_HHmmss", CultureInfo.InvariantCulture);
    string safeTitle = MakeSafeFileName(GetDocTitleOrName(doc));
    string outPath = Path.Combine(outputFolder, safeTitle + "_SW_Report_" + stamp + ".xlsx");

    ReportPayload payload = CollectActiveDoc(swApp, doc, log);

    WriteXlsxOpenXml(outPath, payload);

    return outPath;
}

        /// <summary>Imports an exported .xlsx report and applies PROP_* columns as custom properties to referenced documents.</summary>
        public static void ImportExcelAndApplyProperties(ISldWorks swApp, string xlsxPath, Action<string> log)
{
    if (swApp == null) throw new ArgumentNullException("swApp");
    if (!File.Exists(xlsxPath)) throw new FileNotFoundException("Excel file not found.", xlsxPath);

    OpenXmlTable table = ReadComponentsSheetOpenXml(xlsxPath);

    if (!table.Headers.ContainsKey("FilePath") || !table.Headers.ContainsKey("Config"))
        throw new InvalidOperationException("Missing required columns: FilePath and/or Config.");

    List<string> propCols = table.Headers.Keys
        .Where(h => h.StartsWith("PROP_", StringComparison.OrdinalIgnoreCase))
        .ToList();

    if (propCols.Count == 0)
    {
        if (log != null) log("No PROP_* columns found. Nothing to import.\r\n");
        return;
    }

    int applied = 0;
    int skipped = 0;

    foreach (string[] r in table.Data)
    {
        string filePath = GetCell(r, table.Headers["FilePath"]).Trim();
        string cfg = GetCell(r, table.Headers["Config"]).Trim();

        if (string.IsNullOrWhiteSpace(filePath) || !File.Exists(filePath))
        {
            skipped++;
            continue;
        }

        ModelDoc2 model = null;
        try
        {
            model = EnsureModelOpen(swApp, filePath, log);
            if (model == null)
            {
                skipped++;
                continue;
            }

            IModelDocExtension ext = model.Extension;
            CustomPropertyManager cpm = ext.CustomPropertyManager[cfg ?? ""];

            foreach (string h in propCols)
            {
                int idx = table.Headers[h];
                string propName = h.Substring("PROP_".Length);
                string newValue = GetCell(r, idx).Trim();

                if (propName.Length == 0) continue;
                if (newValue.Length == 0) continue; // skip empties; change if you want to clear

                int res = cpm.Set2(propName, newValue);
                if (res == 0)
                {
                    cpm.Add3(propName, (int)swCustomInfoType_e.swCustomInfoText, newValue,
                        (int)swCustomPropertyAddOption_e.swCustomPropertyReplaceValue);
                }
            }

            applied++;
        }
        catch (Exception ex)
        {
            if (log != null) log("Import failed for '" + filePath + "': " + ex.Message + "\r\n");
            skipped++;
        }
        finally
        {
            if (model != null && Marshal.IsComObject(model))
                Marshal.FinalReleaseComObject(model);
        }
    }

    if (log != null) log("Import complete. Applied: " + applied + ", Skipped: " + skipped + "\r\n");
}

        // -------------------------
        // Data
        // -------------------------
        private sealed class ReportPayload
        {
            public string DocType = "";
            public string Title = "";
            public string Path = "";
            public string ActiveConfig = "";
            public string Units = "";
            public string CreatedAt = "";

            public AssemblySummary Assembly = new AssemblySummary();
            public DrawingSummary Drawing = new DrawingSummary();

            public List<ComponentRow> Components = new List<ComponentRow>();
        }

        private sealed class AssemblySummary
        {
            public int ComponentCount;
            public int UniqueFileCount;
            public int SuppressedCount;
            public int MateCountApprox;
            public string InterferenceStatus = "n/a";
            public int InterferenceCount = 0;
        }

        private sealed class DrawingSummary
        {
            public string SheetScale = "n/a";
            public int ViewCount = 0;
        }

        private sealed class ComponentRow
        {
            public string ComponentPath = "";
            public string FilePath = "";
            public string DocType = "";
            public string Config = "";
            public bool Suppressed;

            public string Units = "";
            public string MaterialName = "";
            public string MaterialDb = "";
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

            public string FeatureCounts = "";

            // editable placeholders
            public string PROP_PartNo = "";
            public string PROP_Description = "";
            public string PROP_Revision = "";
            public string PROP_Finish = "";
        }

        // -------------------------
        // Collect
        // -------------------------
        private static ReportPayload CollectActiveDoc(ISldWorks swApp, ModelDoc2 doc, Action<string> log)
        {
            var payload = new ReportPayload();
            payload.DocType = GetDocTypeName(doc);
            payload.Title = GetDocTitleOrName(doc);
            payload.Path = doc.GetPathName() ?? "";
            payload.ActiveConfig = doc.ConfigurationManager != null && doc.ConfigurationManager.ActiveConfiguration != null
                ? (doc.ConfigurationManager.ActiveConfiguration.Name ?? "")
                : "";
            payload.Units = TryGetUnitsString(doc);
            payload.CreatedAt = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture);

            int type = doc.GetType();
            if (type == (int)swDocumentTypes_e.swDocASSEMBLY)
            {
                CollectAssembly(swApp, (AssemblyDoc)doc, payload, log);
            }
            else if (type == (int)swDocumentTypes_e.swDocPART)
            {
                payload.Components.Add(CollectModelRow(swApp, doc, "(ActivePart)", false, payload.ActiveConfig, log));
            }
            else if (type == (int)swDocumentTypes_e.swDocDRAWING)
            {
                CollectDrawing((DrawingDoc)doc, payload, log);
            }

            return payload;
        }

        private static void CollectAssembly(ISldWorks swApp, AssemblyDoc assy, ReportPayload payload, Action<string> log)
        {
            object compsObj = assy.GetComponents(true);
            object[] compsArr = compsObj as object[] ?? new object[0];

            var comps = compsArr.OfType<Component2>().ToList();

            payload.Assembly.ComponentCount = comps.Count;
            payload.Assembly.SuppressedCount = comps.Count(c => c.IsSuppressed());
            payload.Assembly.UniqueFileCount = comps.Select(c => (c.GetPathName() ?? ""))
                .Where(p => p.Length > 0)
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .Count();

            payload.Assembly.MateCountApprox = CountMatesApprox((ModelDoc2)assy);

            TryInterferenceDetectionDynamic(assy, payload, log);

            foreach (var c in comps)
            {
                string compPath = BuildComponentPath(c);
                string cfg = c.ReferencedConfiguration ?? "";
                bool suppressed = c.IsSuppressed();

                ModelDoc2 cDoc = null;
                try
                {
                    cDoc = c.GetModelDoc2() as ModelDoc2;
                    if (cDoc == null)
                    {
                        payload.Components.Add(new ComponentRow
                        {
                            ComponentPath = compPath,
                            FilePath = c.GetPathName() ?? "",
                            DocType = "Unresolved",
                            Config = cfg,
                            Suppressed = suppressed,
                            Units = payload.Units
                        });
                        continue;
                    }

                    ComponentRow row = CollectModelRow(swApp, cDoc, compPath, suppressed, cfg, log);
                    row.FilePath = c.GetPathName() ?? row.FilePath;
                    payload.Components.Add(row);
                }
                catch (Exception ex)
                {
                    if (log != null) log("Component read error '" + compPath + "': " + ex.Message + "\r\n");
                }
                finally
                {
                    if (cDoc != null && Marshal.IsComObject(cDoc))
                        Marshal.FinalReleaseComObject(cDoc);
                }
            }
        }

        private static void CollectDrawing(DrawingDoc drw, ReportPayload payload, Action<string> log)
        {
            try
            {
                object sheetObj = drw.GetCurrentSheet();
                if (sheetObj != null)
                {
                    double num = 0, den = 0;

                    // Use dynamic to avoid interop signature mismatch
                    dynamic dSheet = sheetObj;
                    dSheet.GetScale(out num, out den);

                    payload.Drawing.SheetScale =
                        num.ToString(CultureInfo.InvariantCulture) + ":" +
                        den.ToString(CultureInfo.InvariantCulture);
                }
            }
            catch
            {
                // leave as "n/a"
            }
        }

        private static ComponentRow CollectModelRow(ISldWorks swApp, ModelDoc2 model, string componentPath, bool suppressed, string config, Action<string> log)
        {
            var row = new ComponentRow();
            row.ComponentPath = componentPath;
            row.FilePath = model.GetPathName() ?? "";
            row.DocType = GetDocTypeName(model);
            row.Config = string.IsNullOrWhiteSpace(config)
                ? (model.ConfigurationManager != null && model.ConfigurationManager.ActiveConfiguration != null ? (model.ConfigurationManager.ActiveConfiguration.Name ?? "") : "")
                : config;
            row.Suppressed = suppressed;
            row.Units = TryGetUnitsString(model);

            // Material (older interop)
            string db, mat;
            TryGetMaterialCompat(model, row.Config, out db, out mat);
            row.MaterialDb = db ?? "";
            row.MaterialName = mat ?? "";

            // Mass props (interop-safe)
            TryMassPropsCompat(model, row);

            // Bounding box (part only reliable)
            TryBoundingBoxPart(model, row);

            // Feature counts (cast-safe)
            row.FeatureCounts = GetFeatureCountsCompat(model);

            // preload some editable properties
            TryReadExistingPropsCompat(model, row.Config, row);

            return row;
        }

        // -------------------------
        // Excel write
        // -------------------------
        // removed ClosedXML helper: ReadHeaders


        // removed ClosedXML helper: ReadHeaders


        private static object NaNtoEmpty(double v)
        {
            if (double.IsNaN(v)) return "";
            return v;
        }


        // -------------------------
        // Compat SW helpers
        // -------------------------
        private static void TryGetMaterialCompat(ModelDoc2 doc, string configName, out string db, out string matName)
        {
            db = "";
            matName = "";

            try
            {
                // Use dynamic to avoid interop signature differences.
                dynamic ext = doc.Extension;

                // Different SW versions expose different methods:
                // - GetMaterialPropertyName2(config, out db, out mat)
                // - GetMaterialPropertyName(config, out db, out mat)
                try
                {
                    ext.GetMaterialPropertyName2(configName, out db, out matName);
                    return;
                }
                catch { }

                try
                {
                    ext.GetMaterialPropertyName(configName, out db, out matName);
                    return;
                }
                catch { }

                // Some environments expose material via other APIs; if none available, leave blank.
            }
            catch
            {
                // leave blank
            }
        }


        private static void TryMassPropsCompat(ModelDoc2 doc, ComponentRow row)
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

                // Inertia: try via dynamic if present; otherwise leave blank
                try
                {
                    dynamic d = mp;
                    double[] p = d.PrincipalMomentsOfInertia;
                    if (p != null && p.Length >= 3)
                    {
                        row.Ixx = p[0]; row.Iyy = p[1]; row.Izz = p[2];
                    }
                }
                catch { }
            }
            catch { }
        }

        private static void TryBoundingBoxPart(ModelDoc2 doc, ComponentRow row)
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

        private static string GetFeatureCountsCompat(ModelDoc2 doc)
        {
            try
            {
                var counts = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);

                Feature f = doc.FirstFeature() as Feature;
                while (f != null)
                {
                    string t = f.GetTypeName2() ?? "";
                    if (t.Length > 0)
                    {
                        int cur;
                        counts.TryGetValue(t, out cur);
                        counts[t] = cur + 1;
                    }
                    f = f.GetNextFeature() as Feature;
                }

                return string.Join(";", counts
                    .OrderByDescending(kv => kv.Value)
                    .ThenBy(kv => kv.Key)
                    .Select(kv => kv.Key + "=" + kv.Value));
            }
            catch
            {
                return "";
            }
        }

        private static void TryReadExistingPropsCompat(ModelDoc2 doc, string cfg, ComponentRow row)
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
                // Older signature: Get4(name, useCached, out valOut, out resolved)
                cpm.Get4(name, false, out valOut, out resolved);
                return !string.IsNullOrWhiteSpace(resolved) ? resolved : (valOut ?? "");
            }
            catch
            {
                return "";
            }
        }

        private static int CountMatesApprox(ModelDoc2 asmDoc)
        {
            try
            {
                int mates = 0;
                Feature f = asmDoc.FirstFeature() as Feature;
                while (f != null)
                {
                    string t = f.GetTypeName2() ?? "";
                    if (string.Equals(t, "MateGroup", StringComparison.OrdinalIgnoreCase))
                    {
                        Feature sub = f.GetFirstSubFeature() as Feature;
                        while (sub != null)
                        {
                            string st = sub.GetTypeName2() ?? "";
                            if (st.IndexOf("Mate", StringComparison.OrdinalIgnoreCase) >= 0) mates++;
                            sub = sub.GetNextSubFeature() as Feature;
                        }
                    }
                    f = f.GetNextFeature() as Feature;
                }
                return mates;
            }
            catch
            {
                return 0;
            }
        }

        private static void TryInterferenceDetectionDynamic(AssemblyDoc assy, ReportPayload payload, Action<string> log)
        {
            try
            {
                // Interop-safe: use dynamic COM call
                dynamic dAssy = assy;
                object mgrObj = dAssy.GetInterferenceDetectionManager();
                if (mgrObj == null)
                {
                    payload.Assembly.InterferenceStatus = "n/a";
                    return;
                }

                dynamic mgr = mgrObj;
                mgr.TreatCoincidenceAsInterference = false;
                mgr.MakeInterferingPartsTransparent = false;
                mgr.IncludeMultibodyPartInterferences = true;

                object resultsObj = mgr.GetInterferences();
                object[] results = resultsObj as object[];
                int count = results != null ? results.Length : 0;

                payload.Assembly.InterferenceCount = count;
                payload.Assembly.InterferenceStatus = count > 0 ? "FAIL" : "PASS";

                if (results != null)
                {
                    for (int i = 0; i < results.Length; i++)
                    {
                        object o = results[i];
                        try { if (o != null && Marshal.IsComObject(o)) Marshal.FinalReleaseComObject(o); } catch { }
                    }
                }

                try { if (mgrObj != null && Marshal.IsComObject(mgrObj)) Marshal.FinalReleaseComObject(mgrObj); } catch { }
            }
            catch
            {
                payload.Assembly.InterferenceStatus = "n/a";
            }
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

        // -------------------------
        // Import helpers
        // -------------------------
        private static ModelDoc2 EnsureModelOpen(ISldWorks swApp, string filePath, Action<string> log)
        {
            try
            {
                ModelDoc2 openDoc = swApp.GetOpenDocumentByName(filePath) as ModelDoc2;
                if (openDoc != null) return openDoc;
            }
            catch { }

            int errs = 0, warns = 0;
            int docType = GuessDocTypeFromExtension(filePath);

            ModelDoc2 model = swApp.OpenDoc6(
                filePath,
                docType,
                (int)swOpenDocOptions_e.swOpenDocOptions_Silent,
                "",
                ref errs,
                ref warns) as ModelDoc2;

            if (model == null && log != null)
                log("Failed to open: " + filePath + " (err=" + errs + ", warn=" + warns + ")\r\n");

            return model;
        }

        private static int GuessDocTypeFromExtension(string path)
        {
            string ext = (Path.GetExtension(path) ?? "").ToLowerInvariant();
            switch (ext)
            {
                case ".sldprt": return (int)swDocumentTypes_e.swDocPART;
                case ".sldasm": return (int)swDocumentTypes_e.swDocASSEMBLY;
                case ".slddrw": return (int)swDocumentTypes_e.swDocDRAWING;
                default: return (int)swDocumentTypes_e.swDocPART;
            }
        }

        // removed ClosedXML helper: ReadHeaders


        // removed ClosedXML helper: FindColumn


        // -------------------------
        // Misc
        // -------------------------
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
            int t = doc.GetType();
            if (t == (int)swDocumentTypes_e.swDocPART) return "Part";
            if (t == (int)swDocumentTypes_e.swDocASSEMBLY) return "Assembly";
            if (t == (int)swDocumentTypes_e.swDocDRAWING) return "Drawing";
            return "Unknown";
        }

        private static string GetDocTitleOrName(ModelDoc2 doc)
        {
            try
            {
                string p = doc.GetPathName();
                if (!string.IsNullOrWhiteSpace(p)) return Path.GetFileNameWithoutExtension(p);
                return doc.GetTitle() ?? "Untitled";
            }
            catch { return "Untitled"; }
        }

        private static string MakeSafeFileName(string s)
        {
            foreach (char ch in Path.GetInvalidFileNameChars())
                s = s.Replace(ch, '_');
            return s;
        }

// -------------------------
// OpenXML XLSX (no Excel install)
// -------------------------
private static void WriteXlsxOpenXml(string filePath, ReportPayload payload)
{
    if (File.Exists(filePath)) File.Delete(filePath);

    using (SpreadsheetDocument document = SpreadsheetDocument.Create(filePath, SpreadsheetDocumentType.Workbook))
    {
        WorkbookPart wbPart = document.AddWorkbookPart();
        wbPart.Workbook = new Workbook();

        WorksheetPart wsSummaryPart = wbPart.AddNewPart<WorksheetPart>();
        WorksheetPart wsComponentsPart = wbPart.AddNewPart<WorksheetPart>();

        wsSummaryPart.Worksheet = new Worksheet(new SheetData());
        wsComponentsPart.Worksheet = new Worksheet(new SheetData());

        Sheets sheets = wbPart.Workbook.AppendChild(new Sheets());

        sheets.Append(new OxSheet { Id = wbPart.GetIdOfPart(wsSummaryPart), SheetId = 1U, Name = "Summary" });
        sheets.Append(new OxSheet { Id = wbPart.GetIdOfPart(wsComponentsPart), SheetId = 2U, Name = "Components" });

        FillSummarySheet(wsSummaryPart.Worksheet.GetFirstChild<SheetData>(), payload);
        FillComponentsSheet(wsComponentsPart.Worksheet.GetFirstChild<SheetData>(), payload);

        wsSummaryPart.Worksheet.Save();
        wsComponentsPart.Worksheet.Save();
        wbPart.Workbook.Save();
    }
}

private static void FillSummarySheet(SheetData sheetData, ReportPayload payload)
{
    uint r = 1U;

    AppendRow(sheetData, r++, "CreatedAt", payload.CreatedAt);
    AppendRow(sheetData, r++, "DocType", payload.DocType);
    AppendRow(sheetData, r++, "Title", payload.Title);
    AppendRow(sheetData, r++, "Path", payload.Path);
    AppendRow(sheetData, r++, "ActiveConfig", payload.ActiveConfig);
    AppendRow(sheetData, r++, "Units", payload.Units);

    r++;
    AppendRow(sheetData, r++, "Assembly.ComponentCount", payload.Assembly.ComponentCount);
    AppendRow(sheetData, r++, "Assembly.UniqueFileCount", payload.Assembly.UniqueFileCount);
    AppendRow(sheetData, r++, "Assembly.SuppressedCount", payload.Assembly.SuppressedCount);
    AppendRow(sheetData, r++, "Assembly.MateCountApprox", payload.Assembly.MateCountApprox);
    AppendRow(sheetData, r++, "Assembly.InterferenceStatus", payload.Assembly.InterferenceStatus);
    AppendRow(sheetData, r++, "Assembly.InterferenceCount", payload.Assembly.InterferenceCount);

    r++;
    AppendRow(sheetData, r++, "Drawing.SheetScale", payload.Drawing.SheetScale);
    AppendRow(sheetData, r++, "Drawing.ViewCount", payload.Drawing.ViewCount);
}

private static void FillComponentsSheet(SheetData sheetData, ReportPayload payload)
{
    string[] headers = new[]
    {
        "ComponentPath","FilePath","DocType","Config","Suppressed","Units",
        "MaterialDb","MaterialName","Density",
        "Mass","CoG_X","CoG_Y","CoG_Z",
        "Ixx","Iyy","Izz",
        "Volume","SurfaceArea",
        "BBX_X","BBX_Y","BBX_Z",
        "FeatureCounts",
        "PROP_PartNo","PROP_Description","PROP_Revision","PROP_Finish"
    };

    uint r = 1U;
    AppendRow(sheetData, r++, headers);

    foreach (var row in payload.Components)
    {
        object[] cells = new object[]
        {
            row.ComponentPath,
            row.FilePath,
            row.DocType,
            row.Config,
            row.Suppressed ? 1 : 0,
            row.Units,

            row.MaterialDb,
            row.MaterialName,
            NaNtoEmpty(row.Density),

            NaNtoEmpty(row.Mass),
            NaNtoEmpty(row.CoG_X),
            NaNtoEmpty(row.CoG_Y),
            NaNtoEmpty(row.CoG_Z),

            NaNtoEmpty(row.Ixx),
            NaNtoEmpty(row.Iyy),
            NaNtoEmpty(row.Izz),

            NaNtoEmpty(row.Volume),
            NaNtoEmpty(row.SurfaceArea),

            NaNtoEmpty(row.BBX_X),
            NaNtoEmpty(row.BBX_Y),
            NaNtoEmpty(row.BBX_Z),

            row.FeatureCounts,

            row.PROP_PartNo,
            row.PROP_Description,
            row.PROP_Revision,
            row.PROP_Finish
        };

        AppendRow(sheetData, r++, cells);
    }
}

private static void AppendRow(SheetData sheetData, uint rowIndex, string key, object value)
{
    Row row = new Row { RowIndex = rowIndex };
    row.Append(CreateCell("A", rowIndex, key));
    row.Append(CreateCell("B", rowIndex, value));
    sheetData.Append(row);
}

private static void AppendRow(SheetData sheetData, uint rowIndex, params string[] values)
{
    Row row = new Row { RowIndex = rowIndex };
    for (int i = 0; i < values.Length; i++)
    {
        string col = ColumnName(i + 1);
        row.Append(CreateCell(col, rowIndex, values[i]));
    }
    sheetData.Append(row);
}

private static void AppendRow(SheetData sheetData, uint rowIndex, params object[] values)
{
    Row row = new Row { RowIndex = rowIndex };
    for (int i = 0; i < values.Length; i++)
    {
        string col = ColumnName(i + 1);
        row.Append(CreateCell(col, rowIndex, values[i]));
    }
    sheetData.Append(row);
}

private static Cell CreateCell(string columnName, uint rowIndex, object value)
{
    Cell cell = new Cell();
    cell.CellReference = columnName + rowIndex.ToString(CultureInfo.InvariantCulture);

    if (value == null)
    {
        cell.DataType = CellValues.String;
        cell.CellValue = new CellValue("");
        return cell;
    }

    // numeric -> store as number
    if (value is int || value is long || value is float || value is double || value is decimal)
    {
        double d;
        if (double.TryParse(Convert.ToString(value, CultureInfo.InvariantCulture), NumberStyles.Any, CultureInfo.InvariantCulture, out d))
        {
            cell.DataType = CellValues.Number;
            cell.CellValue = new CellValue(d.ToString(CultureInfo.InvariantCulture));
            return cell;
        }
    }

    cell.DataType = CellValues.String;
    cell.CellValue = new CellValue(Convert.ToString(value, CultureInfo.InvariantCulture) ?? "");
    return cell;
}

private static string ColumnName(int index1Based)
{
    int dividend = index1Based;
    string colName = "";
    while (dividend > 0)
    {
        int modulo = (dividend - 1) % 26;
        colName = Convert.ToChar('A' + modulo) + colName;
        dividend = (dividend - modulo) / 26;
    }
    return colName;
}

private sealed class OpenXmlTable
{
    public Dictionary<string, int> Headers = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
    public List<string[]> Data = new List<string[]>();
}

private static OpenXmlTable ReadComponentsSheetOpenXml(string xlsxPath)
{
    OpenXmlTable table = new OpenXmlTable();

    using (SpreadsheetDocument doc = SpreadsheetDocument.Open(xlsxPath, false))
    {
        WorkbookPart wbPart = doc.WorkbookPart;
        if (wbPart == null) throw new InvalidOperationException("Invalid workbook.");

        OxSheet sheet = wbPart.Workbook.Sheets.Elements<OxSheet>()
            .FirstOrDefault(s => string.Equals(s.Name, "Components", StringComparison.OrdinalIgnoreCase));

        if (sheet == null) throw new InvalidOperationException("Worksheet 'Components' not found.");

        WorksheetPart wsPart = (WorksheetPart)wbPart.GetPartById(sheet.Id);
        SheetData sheetData = wsPart.Worksheet.GetFirstChild<SheetData>();
        if (sheetData == null) throw new InvalidOperationException("No sheet data.");

        List<Row> rows = sheetData.Elements<Row>().ToList();
        if (rows.Count == 0) throw new InvalidOperationException("No rows found.");

        string[] headerCells = ReadRowAsArray(doc, rows[0]);
        for (int i = 0; i < headerCells.Length; i++)
        {
            string h = (headerCells[i] ?? "").Trim();
            if (h.Length > 0 && !table.Headers.ContainsKey(h))
                table.Headers[h] = i;
        }

        for (int i = 1; i < rows.Count; i++)
        {
            table.Data.Add(ReadRowAsArray(doc, rows[i], headerCells.Length));
        }
    }

    return table;
}

private static string[] ReadRowAsArray(SpreadsheetDocument doc, Row row, int minLen)
{
    List<Cell> cells = row.Elements<Cell>().ToList();

    int maxCol = 0;
    for (int i = 0; i < cells.Count; i++)
    {
        int col = GetColumnIndexFromCellReference(cells[i].CellReference);
        if (col > maxCol) maxCol = col;
    }

    int len = Math.Max(maxCol, minLen);
    if (len <= 0) len = minLen > 0 ? minLen : 1;

    string[] arr = new string[len];
    for (int i = 0; i < arr.Length; i++) arr[i] = "";

    for (int i = 0; i < cells.Count; i++)
    {
        Cell c = cells[i];
        int colIndex1 = GetColumnIndexFromCellReference(c.CellReference);
        if (colIndex1 <= 0) continue;
        int idx = colIndex1 - 1;
        if (idx >= arr.Length) continue;
        arr[idx] = GetCellValue(doc, c);
    }

    return arr;
}

private static string[] ReadRowAsArray(SpreadsheetDocument doc, Row row)
{
    return ReadRowAsArray(doc, row, 0);
}

private static string GetCell(string[] row, int idx)
{
    if (row == null) return "";
    if (idx < 0 || idx >= row.Length) return "";
    return row[idx] ?? "";
}

private static string GetCellValue(SpreadsheetDocument doc, Cell cell)
{
    if (cell == null) return "";
    string value = cell.CellValue != null ? cell.CellValue.InnerText : "";
    if (cell.DataType == null) return value ?? "";

    if (cell.DataType.Value == CellValues.SharedString)
    {
        SharedStringTablePart sst = doc.WorkbookPart.SharedStringTablePart;
        if (sst == null) return value ?? "";
        int i;
        if (int.TryParse(value, out i))
        {
            return sst.SharedStringTable.Elements<SharedStringItem>().ElementAt(i).InnerText ?? "";
        }
    }

    return value ?? "";
}

private static int GetColumnIndexFromCellReference(StringValue cellRef)
{
    if (cellRef == null) return 0;
    string reference = cellRef.Value;
    if (string.IsNullOrEmpty(reference)) return 0;

    int i = 0;
    while (i < reference.Length && char.IsLetter(reference[i])) i++;
    string colLetters = reference.Substring(0, i).ToUpperInvariant();

    int sum = 0;
    for (int j = 0; j < colLetters.Length; j++)
    {
        sum *= 26;
        sum += (colLetters[j] - 'A' + 1);
    }
    return sum;
}


    }
}
