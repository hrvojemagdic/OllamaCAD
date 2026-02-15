using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using SolidWorks.Interop.sldworks;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace OllamaCAD
{   
    /// <summary>
    /// Manages per-project memory storage for the OllamaCAD add-in.
    ///
    /// Responsibilities:
    /// - Creates and maintains a memory folder tied to the active SOLIDWORKS document.
    /// - Stores conversation history (JSONL), summary text, settings, and screenshots.
    /// - Provides thread-safe async read/write access using a SemaphoreSlim gate.
    /// - Supports clearing memory and opening the project folder in Explorer.
    ///
    /// Each CAD file gets its own ".ollama" memory folder, enabling isolated,
    /// persistent AI context per project.
    /// </summary>
    internal sealed class ProjectMemoryStore
    {
        private readonly SemaphoreSlim _gate = new SemaphoreSlim(1, 1);

        public string RootFolder { get; private set; }
        public string ScreenshotsFolder => Path.Combine(RootFolder, "screenshots");
        public string SettingsPath => Path.Combine(RootFolder, "settings.json");
        public string SummaryPath => Path.Combine(RootFolder, "summary.txt");
        public string ConversationPath => Path.Combine(RootFolder, "conversation.jsonl");

        public async Task EnsureForActiveDocAsync(ISldWorks swApp)
        {
            string root = ResolveRootFolder(swApp);

            if (string.Equals(root, RootFolder, StringComparison.OrdinalIgnoreCase))
                return;

            RootFolder = root;

            Directory.CreateDirectory(RootFolder);
            Directory.CreateDirectory(ScreenshotsFolder);

            // Ensure files exist
            if (!File.Exists(SummaryPath))
                await IoCompat.WriteAllTextAsync(SummaryPath, "", Encoding.UTF8);

            if (!File.Exists(ConversationPath))
                await IoCompat.WriteAllTextAsync(ConversationPath, "", Encoding.UTF8);

            if (!File.Exists(SettingsPath))
            {
                var s = new MemorySettings();
                await SaveSettingsAsync(s);
            }
        }

        public async Task<MemorySettings> LoadSettingsAsync()
        {
            await _gate.WaitAsync();
            try
            {
                if (!File.Exists(SettingsPath))
                    return new MemorySettings();

                string json = await IoCompat.ReadAllTextAsync(SettingsPath, Encoding.UTF8);
                var s = JsonConvert.DeserializeObject<MemorySettings>(json);
                return s ?? new MemorySettings();
            }
            finally { _gate.Release(); }
        }

        public async Task SaveSettingsAsync(MemorySettings settings)
        {
            await _gate.WaitAsync();
            try
            {
                string json = JsonConvert.SerializeObject(settings, Formatting.Indented);
                await IoCompat.WriteAllTextAsync(SettingsPath, json, Encoding.UTF8);
            }
            finally { _gate.Release(); }
        }

        public async Task<string> LoadSummaryAsync()
        {
            await _gate.WaitAsync();
            try
            {
                if (!File.Exists(SummaryPath)) return "";
                return await IoCompat.ReadAllTextAsync(SummaryPath, Encoding.UTF8);
            }
            finally { _gate.Release(); }
        }

        public async Task AppendTurnAsync(ChatTurn turn)
        {
            await _gate.WaitAsync();
            try
            {
                string line = JsonConvert.SerializeObject(turn, Formatting.None);
                await IoCompat.AppendAllTextAsync(ConversationPath, line + System.Environment.NewLine, Encoding.UTF8);
            }
            finally { _gate.Release(); }
        }

        public async Task<List<ChatTurn>> LoadRecentTurnsAsync(int maxTurns)
        {
            await _gate.WaitAsync();
            try
            {
                if (!File.Exists(ConversationPath)) return new List<ChatTurn>();

                var lines = await IoCompat.ReadAllLinesAsync(ConversationPath, Encoding.UTF8);
                var recent = lines
                    .Where(l => !string.IsNullOrWhiteSpace(l))
                    .Reverse()
                    .Take(Math.Max(1, maxTurns))
                    .Reverse()
                    .Select(l =>
                    {
                        try { return JsonConvert.DeserializeObject<ChatTurn>(l); }
                        catch { return null; }
                    })
                    .Where(t => t != null)
                    .ToList();

                return recent;
            }
            finally { _gate.Release(); }
        }

        public async Task ClearAsync()
        {
            await _gate.WaitAsync();
            try
            {
                if (File.Exists(ConversationPath)) File.WriteAllText(ConversationPath, "");
                if (File.Exists(SummaryPath)) File.WriteAllText(SummaryPath, "");
                if (Directory.Exists(ScreenshotsFolder))
                {
                    foreach (var f in Directory.GetFiles(ScreenshotsFolder))
                    {
                        try { File.Delete(f); } catch { }
                    }
                }
            }
            finally { _gate.Release(); }
        }

        public void OpenFolderInExplorer()
        {
            if (string.IsNullOrWhiteSpace(RootFolder) || !Directory.Exists(RootFolder))
                return;

            Process.Start(new ProcessStartInfo
            {
                FileName = "explorer.exe",
                Arguments = $"\"{RootFolder}\"",
                UseShellExecute = true
            });
        }

        public string SaveScreenshotPng(byte[] pngBytes, string prefix)
        {
            Directory.CreateDirectory(ScreenshotsFolder);

            string safePrefix = string.IsNullOrWhiteSpace(prefix) ? "shot" : prefix;
            string name = $"{DateTime.Now:yyyy-MM-dd_HH-mm-ss}_{safePrefix}.png";
            string path = Path.Combine(ScreenshotsFolder, name);

            File.WriteAllBytes(path, pngBytes);
            return path;
        }

        private static string ResolveRootFolder(ISldWorks swApp)
        {
            try
            {
                var doc = (ModelDoc2)swApp.ActiveDoc;
                if (doc != null)
                {
                    string p = doc.GetPathName();
                    if (!string.IsNullOrWhiteSpace(p) && File.Exists(p))
                    {
                        // Folder next to CAD file
                        return p + ".ollama";
                    }

                    // Unsaved document fallback by title
                    string title = doc.GetTitle();
                    return Path.Combine(
                        System.Environment.GetFolderPath(System.Environment.SpecialFolder.LocalApplicationData),
                        "OllamaCAD",
                        "Unsaved",
                        SanitizeFileName(title)
                    );
                }
            }
            catch { }

            return Path.Combine(
                System.Environment.GetFolderPath(System.Environment.SpecialFolder.LocalApplicationData),
                "OllamaCAD",
                "UnknownDoc"
            );
        }

        private static string SanitizeFileName(string s)
        {
            if (string.IsNullOrWhiteSpace(s)) return "Untitled";
            foreach (char c in Path.GetInvalidFileNameChars())
                s = s.Replace(c, '_');
            return s.Trim();
        }
    }

    /// <summary>
    /// Configuration object for project memory and RAG behavior.
    ///
    /// Includes:
    /// - Chat behavior settings (memory enabled, screenshots, summarization frequency).
    /// - Selected Ollama model for chat.
    /// - RAG configuration (folders, Python path, Poppler path, Top-K).
    /// - Model selections for OCR, QA, and embeddings.
    ///
    /// Serialized to settings.json inside the project memory folder.
    /// </summary>
    class MemorySettings
    {
        // Memory/chat
        public bool EnableMemory { get; set; } = true;
        public bool IncludeScreenshotInPrompt { get; set; } = false;
        public bool SaveScreenshotsToDisk { get; set; } = true;

        public int MaxRecentTurns { get; set; } = 6;
        public int SummarizeEveryNTurns { get; set; } = 12;

        // Chat model
        public string ModelName { get; set; } = "qwen3-vl:8b-instruct-q4_K_M";

        // RAG
        public bool EnableRagOnly { get; set; } = false;

        public string RagFolderName { get; set; } = "OllamaRAG";
        public string RagStoreFolderName { get; set; } = "rag_store";

        public string PythonExePath { get; set; } = "";
        public string RagScriptPath { get; set; } = "";   // full path only, otherwise ignored by RagRunner

        public string PopplerBinPath { get; set; } = "";  // if empty -> env var OLLAMACAD_POPPLER

        public int RagTopK { get; set; } = 10;

        // RAG model selection
        public string OcrModelName { get; set; } = "qwen3-vl:8b-instruct-q4_K_M";
        public string QaModelName { get; set; } = "gemma3:12b-it-q4_K_M";
        public string EmbedModelName { get; set; } = "qwen3-embedding:8b-q4_K_M";
    }

    /// <summary>
    /// Represents a single chat interaction stored in conversation.jsonl.
    ///
    /// Contains:
    /// - Timestamp
    /// - Role (user/assistant)
    /// - Message content
    /// - Model used
    /// - Screenshot metadata (if applicable)
    /// - Active SOLIDWORKS document path
    ///
    /// Used for reconstructing recent context and generating summaries.
    /// </summary>
    internal sealed class ChatTurn
    {
        public DateTime Timestamp { get; set; } = DateTime.Now;
        public string Role { get; set; }              // "user" or "assistant"
        public string Content { get; set; }
        public string Model { get; set; }
        public bool IncludedScreenshotInPrompt { get; set; }
        public string ScreenshotPath { get; set; }    // optional
        public string ActiveDocPath { get; set; }     // optional
    }
}
