using System;
using System.Diagnostics;
using System.IO;
using System.Text;
using System.Threading.Tasks;

namespace OllamaCAD
{   
    /// <summary>
    /// Executes the external Python-based RAG pipeline (OllamaRAG_Multi.py) from the SOLIDWORKS add-in.
    ///
    /// Responsibilities:
    /// - Resolves which Python executable to use (project setting, global venv, project venv, or PATH python).
    /// - Resolves the RAG script location (explicit path, global script, or add-in folder fallback).
    /// - Builds/refreshes the project RAG index from documents inside the project's OllamaRAG folder.
    /// - Runs question answering against the built index (RAG-only mode).
    /// - Passes model selections (OCR/QA/embedding), Top-K, and optional Poppler path to the script.
    /// - Captures stdout/stderr and returns structured process results for logging in the UI.
    ///
    /// The RAG index is considered "ready" when faiss.index and meta.pkl exist in the rag_store folder.
    /// </summary>
    internal sealed class RagRunner
    {
        public sealed class ProcResult
        {
            public int ExitCode { get; set; }
            public string Stdout { get; set; }
            public string Stderr { get; set; }
            public bool Ok => ExitCode == 0;
        }

        private static string GlobalDir =>
            Path.Combine(
                System.Environment.GetFolderPath(System.Environment.SpecialFolder.LocalApplicationData),
                "OllamaCAD"
            );

        private static string GlobalVenvPython =>
            Path.Combine(GlobalDir, "venv", "Scripts", "python.exe");

        private static string GlobalScript =>
            Path.Combine(GlobalDir, "OllamaRAG_Multi.py");

        /// <summary>Resolves the Python executable path used to run the RAG script.</summary>
        public static string ResolvePythonExe(string projectRoot, MemorySettings s)
        {
            // 1) explicit setting
            if (!string.IsNullOrWhiteSpace(s.PythonExePath) && File.Exists(s.PythonExePath))
                return s.PythonExePath;

            // 2) global venv (preferred)
            if (File.Exists(GlobalVenvPython))
                return GlobalVenvPython;

            // 3) project venv fallback
            string projVenv = Path.Combine(projectRoot, "venv", "Scripts", "python.exe");
            if (File.Exists(projVenv))
                return projVenv;

            // 4) fallback
            return "python";
        }

        /// <summary>Resolves the RAG script path (OllamaRAG_Multi.py) to execute.</summary>
        public static string ResolveScriptPath(string projectRoot, MemorySettings s)
        {
            // 1) explicit full path
            if (!string.IsNullOrWhiteSpace(s.RagScriptPath) &&
                Path.IsPathRooted(s.RagScriptPath) &&
                File.Exists(s.RagScriptPath))
                return s.RagScriptPath;

            // 2) global script (preferred)
            if (File.Exists(GlobalScript))
                return GlobalScript;

            // 3) last fallback: next to add-in dll (for first-time copy)
            string baseDir = AppDomain.CurrentDomain.BaseDirectory;
            return Path.Combine(baseDir, "OllamaRAG_Multi.py");
        }

        /// <summary>Builds or refreshes the FAISS index for documents in the project's RAG folder.</summary>
        public static async Task<ProcResult> BuildIndexAsync(string projectRoot, MemorySettings s)
        {
            string pythonExe = ResolvePythonExe(projectRoot, s);
            string script = ResolveScriptPath(projectRoot, s);

            if (!File.Exists(script))
            {
                return new ProcResult
                {
                    ExitCode = 2,
                    Stdout = "",
                    Stderr =
                        "RAG script not found.\n" +
                        $"Expected global: {GlobalScript}\n" +
                        "Run 'Setup Global RAG Environment' or place OllamaRAG_Multi.py in that folder."
                };
            }

            string ragFolder = Path.Combine(projectRoot, s.RagFolderName ?? "OllamaRAG");
            string ragStore = Path.Combine(projectRoot, s.RagStoreFolderName ?? "rag_store");

            Directory.CreateDirectory(ragFolder);
            Directory.CreateDirectory(ragStore);

            var args = new StringBuilder();
            args.Append($"\"{script}\" ");
            args.Append($"--dir \"{ragFolder}\" ");
            args.Append($"--store \"{ragStore}\" ");
            args.Append($"--topk {Math.Max(1, s.RagTopK)} ");

            // Pass models
            if (!string.IsNullOrWhiteSpace(s.OcrModelName)) args.Append($"--ocr \"{EscapeArg(s.OcrModelName)}\" ");
            if (!string.IsNullOrWhiteSpace(s.QaModelName)) args.Append($"--qa \"{EscapeArg(s.QaModelName)}\" ");
            if (!string.IsNullOrWhiteSpace(s.EmbedModelName)) args.Append($"--embed \"{EscapeArg(s.EmbedModelName)}\" ");

            string poppler = s.PopplerBinPath;
            if (string.IsNullOrWhiteSpace(poppler))
                poppler = System.Environment.GetEnvironmentVariable("OLLAMACAD_POPPLER");

            if (!string.IsNullOrWhiteSpace(poppler))
                args.Append($"--poppler \"{poppler}\" ");

            return await RunAsync(pythonExe, args.ToString(), projectRoot);
        }

        /// <summary>Runs a RAG query against the existing index and returns the answer and logs.</summary>
        public static async Task<ProcResult> AskAsync(string projectRoot, MemorySettings s, string question)
        {
            string pythonExe = ResolvePythonExe(projectRoot, s);
            string script = ResolveScriptPath(projectRoot, s);

            if (!File.Exists(script))
            {
                return new ProcResult
                {
                    ExitCode = 2,
                    Stdout = "",
                    Stderr =
                        "RAG script not found.\n" +
                        $"Expected global: {GlobalScript}\n" +
                        "Run 'Setup Global RAG Environment' or place OllamaRAG_Multi.py in that folder."
                };
            }

            string ragStore = Path.Combine(projectRoot, s.RagStoreFolderName ?? "rag_store");
            Directory.CreateDirectory(ragStore);

            var args = new StringBuilder();
            args.Append($"\"{script}\" ");
            args.Append($"--ask \"{EscapeArg(question)}\" ");
            args.Append($"--store \"{ragStore}\" ");
            args.Append($"--topk {Math.Max(1, s.RagTopK)} ");

            // Pass models (QA uses these at answer-time; embed used for query vector)
            if (!string.IsNullOrWhiteSpace(s.OcrModelName)) args.Append($"--ocr \"{EscapeArg(s.OcrModelName)}\" ");
            if (!string.IsNullOrWhiteSpace(s.QaModelName)) args.Append($"--qa \"{EscapeArg(s.QaModelName)}\" ");
            if (!string.IsNullOrWhiteSpace(s.EmbedModelName)) args.Append($"--embed \"{EscapeArg(s.EmbedModelName)}\" ");

            string poppler = s.PopplerBinPath;
            if (string.IsNullOrWhiteSpace(poppler))
                poppler = System.Environment.GetEnvironmentVariable("OLLAMACAD_POPPLER");

            if (!string.IsNullOrWhiteSpace(poppler))
                args.Append($"--poppler \"{poppler}\" ");

            return await RunAsync(pythonExe, args.ToString(), projectRoot);
        }

        private static async Task<ProcResult> RunAsync(string exe, string args, string workDir)
        {
            var psi = new ProcessStartInfo
            {
                FileName = exe,
                Arguments = args,
                WorkingDirectory = workDir,
                UseShellExecute = false,
                RedirectStandardOutput = true,
                RedirectStandardError = true,
                CreateNoWindow = true,
                StandardOutputEncoding = Encoding.UTF8,
                StandardErrorEncoding = Encoding.UTF8
            };

            var p = new Process { StartInfo = psi };

            var stdout = new StringBuilder();
            var stderr = new StringBuilder();

            p.OutputDataReceived += (s, e) => { if (e.Data != null) stdout.AppendLine(e.Data); };
            p.ErrorDataReceived += (s, e) => { if (e.Data != null) stderr.AppendLine(e.Data); };

            p.Start();
            p.BeginOutputReadLine();
            p.BeginErrorReadLine();

            await Task.Run(() => p.WaitForExit());

            return new ProcResult
            {
                ExitCode = p.ExitCode,
                Stdout = stdout.ToString(),
                Stderr = stderr.ToString()
            };
        }

        private static string EscapeArg(string s) => (s ?? "").Replace("\"", "\\\"");

        /// <summary>Returns true if the RAG store contains the required index + metadata files.</summary>
        public static bool IsIndexReady(string projectRoot, MemorySettings s)
        {
            string ragStore = Path.Combine(projectRoot, s.RagStoreFolderName ?? "rag_store");
            string idx = Path.Combine(ragStore, "faiss.index");
            string meta = Path.Combine(ragStore, "meta.pkl");
            return File.Exists(idx) && File.Exists(meta);
        }
    }
}
