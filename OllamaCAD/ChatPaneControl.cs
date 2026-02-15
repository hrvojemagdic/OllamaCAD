using SolidWorks.Interop.sldworks;
using System;
using System.Diagnostics;
using System.IO;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace OllamaCAD
{   
    /// <summary>
    /// Task Pane UI control for the OllamaCAD SOLIDWORKS add-in.
    ///
    /// Responsibilities:
    /// - Builds and manages the chat UI (chat log, input, send, model selection).
    /// - Connects to a local Ollama server for chat/vision responses (optional screenshot input).
    /// - Maintains per-document project memory (turn logging + periodic summary updates).
    /// - Watches the active SOLIDWORKS document and reloads settings/memory when it changes.
    /// - Provides an optional RAG-only mode that answers using documents from the project OllamaRAG folder.
    /// - Supports global RAG environment setup (Python venv + dependencies) and index build/refresh.
    /// - Exports/imports an assembly Excel report and applies properties back into SOLIDWORKS.
    /// </summary>
    public class ChatPaneControl : UserControl
    {

        private System.Windows.Forms.Timer _docWatchTimer;
        private string _lastDocKey = "";
        private bool _docRefreshBusy = false;

        private readonly ISldWorks _swApp;
        private readonly OllamaClient _ollama;

        // Memory
        private readonly ProjectMemoryStore _store;
        private readonly ConversationSummarizer _summarizer;
        private MemorySettings _settings;
        private int _turnCounterSinceSummary = 0;

        // UI
        private TextBox txtChat;
        private TextBox txtInput;
        private Button btnSend;
        private Label lblStatus;

        private ComboBox cboModels;
        private Button btnRefreshModels;

        private CheckBox chkEnableMemory;
        private CheckBox chkIncludeScreenshotInPrompt;
        private CheckBox chkSaveScreenshots;
        private Button btnOpenFolder;
        private Button btnClearMemory;
        private Label lblFolder;

        // RAG: always-visible toggle
        private CheckBox chkRagOnly;

        // RAG: panel (only visible when chkRagOnly checked)
        private FlowLayoutPanel pnlRag;
        private Button btnSetupGlobalRag;
        private Button btnOpenRagFolder;
        private Button btnBuildRagIndex;
        private Label lblRagStatus;

        private Label lblOcr;
        private ComboBox cboOcrModel;

        private Label lblQa;
        private ComboBox cboQaModel;

        private Label lblEmbed;
        private ComboBox cboEmbedModel;

        // Excel output-import
        private Button btnExportExcel;
        private Button btnImportExcel;

        public ChatPaneControl(ISldWorks swApp)
        {
            _swApp = swApp;

            _ollama = new OllamaClient("http://localhost:11434/");
            _store = new ProjectMemoryStore();
            _summarizer = new ConversationSummarizer(_ollama);

            BuildUi();
            _docWatchTimer = new System.Windows.Forms.Timer();
            _docWatchTimer.Interval = 800; // ms
            _docWatchTimer.Tick += async (s, e) => await RefreshIfActiveDocChangedAsync();
            _docWatchTimer.Start();

            Log("OllamaCAD ready.\r\n");

            _ = InitializeAsync();
        }

        /// <summary>Polls for active document changes and reloads project settings/memory when it changes.</summary>
        private async Task RefreshIfActiveDocChangedAsync()
        {
            if (_docRefreshBusy) return;

            try
            {
                _docRefreshBusy = true;

                // Build a stable key for active doc (path if saved, otherwise title)
                string key = "";
                try
                {
                    var doc = (ModelDoc2)_swApp.ActiveDoc;
                    if (doc != null)
                    {
                        string p = doc.GetPathName();
                        if (!string.IsNullOrWhiteSpace(p))
                            key = "PATH:" + p;
                        else
                            key = "TITLE:" + doc.GetTitle();
                    }
                }
                catch { }

                if (string.Equals(key, _lastDocKey, StringComparison.OrdinalIgnoreCase))
                    return;

                _lastDocKey = key;

                // Re-bind store to active doc and reload settings
                await _store.EnsureForActiveDocAsync(_swApp);
                _settings = await _store.LoadSettingsAsync();

                // Apply settings to UI + runtime model
                chkEnableMemory.Checked = _settings.EnableMemory;
                chkIncludeScreenshotInPrompt.Checked = _settings.IncludeScreenshotInPrompt;
                chkSaveScreenshots.Checked = _settings.SaveScreenshotsToDisk;

                chkRagOnly.Checked = _settings.EnableRagOnly;

                // Update chat model in runtime and UI selection
                _ollama.Model = _settings.ModelName;
                if (cboModels.Items.Contains(_ollama.Model))
                    cboModels.SelectedItem = _ollama.Model;

                lblFolder.Text = $"Memory folder: {_store.RootFolder}";

                EnsureRagFolders();
                UpdateRagStatus();
                UpdateRagUiVisibility();

                //Log($"[Active document changed] Now using: {_lastDocKey}\r\n");
            }
            catch
            {
                // ignore polling errors
            }
            finally
            {
                _docRefreshBusy = false;
            }
        }

        /// <summary>Initializes UI state, loads settings, loads available Ollama models, and applies RAG selections.</summary>
        private async Task InitializeAsync()
        {
            await _store.EnsureForActiveDocAsync(_swApp);
            _settings = await _store.LoadSettingsAsync();

            // Apply settings -> UI
            chkEnableMemory.Checked = _settings.EnableMemory;
            chkIncludeScreenshotInPrompt.Checked = _settings.IncludeScreenshotInPrompt;
            chkSaveScreenshots.Checked = _settings.SaveScreenshotsToDisk;

            chkRagOnly.Checked = _settings.EnableRagOnly;

            _ollama.Model = _settings.ModelName;
            lblStatus.Text = $"Model: {_ollama.Model}";

            EnsureRagFolders();
            lblFolder.Text = $"Memory folder: {_store.RootFolder}";
            UpdateRagStatus();

            // Populate models
            await LoadModelsAsync(selectModel: _ollama.Model);

            // Apply RAG model dropdown selections after load
            ApplyRagModelSelections();

            // Finally: apply panel visibility
            UpdateRagUiVisibility();
        }

        private void BuildUi()
        {
            Dock = DockStyle.Fill;

            txtChat = new TextBox
            {
                Multiline = true,
                ReadOnly = true,
                ScrollBars = ScrollBars.Vertical,
                Dock = DockStyle.Fill
            };

            lblStatus = new Label { Dock = DockStyle.Bottom, Height = 18, Text = "Idle" };

            txtInput = new TextBox { Dock = DockStyle.Bottom };

            btnSend = new Button { Text = "Send", Dock = DockStyle.Bottom, Height = 28 };
            btnSend.Click += async (s, e) => await SendAsync();

            txtInput.KeyDown += async (s, e) =>
            {
                if (e.KeyCode == Keys.Enter && !e.Shift)
                {
                    e.SuppressKeyPress = true;
                    await SendAsync();
                }
            };

            // Model dropdown + refresh
            cboModels = new ComboBox { Dock = DockStyle.Bottom, DropDownStyle = ComboBoxStyle.DropDownList };
            cboModels.SelectedIndexChanged += async (s, e) =>
            {
                if (_settings == null) return;
                if (cboModels.SelectedItem is string model && !string.IsNullOrWhiteSpace(model))
                {
                    _ollama.Model = model;
                    lblStatus.Text = $"Model: {_ollama.Model}";
                    _settings.ModelName = model;
                    await _store.SaveSettingsAsync(_settings);
                }
            };

            btnRefreshModels = new Button { Text = "Refresh models", Dock = DockStyle.Bottom, Height = 24 };
            btnRefreshModels.Click += async (s, e) => await LoadModelsAsync(selectModel: _ollama.Model);

            // Memory controls
            chkEnableMemory = new CheckBox { Text = "Enable project memory", Dock = DockStyle.Bottom, Checked = true };
            chkEnableMemory.CheckedChanged += async (s, e) =>
            {
                if (_settings == null) return;
                _settings.EnableMemory = chkEnableMemory.Checked;
                await _store.SaveSettingsAsync(_settings);
            };

            chkIncludeScreenshotInPrompt = new CheckBox { Text = "Include screenshot in prompt", Dock = DockStyle.Bottom, Checked = false };
            chkIncludeScreenshotInPrompt.CheckedChanged += async (s, e) =>
            {
                if (_settings == null) return;
                _settings.IncludeScreenshotInPrompt = chkIncludeScreenshotInPrompt.Checked;
                await _store.SaveSettingsAsync(_settings);
            };

            chkSaveScreenshots = new CheckBox { Text = "Save screenshots to project folder", Dock = DockStyle.Bottom, Checked = true };
            chkSaveScreenshots.CheckedChanged += async (s, e) =>
            {
                if (_settings == null) return;
                _settings.SaveScreenshotsToDisk = chkSaveScreenshots.Checked;
                await _store.SaveSettingsAsync(_settings);
            };

            btnOpenFolder = new Button { Text = "Open memory folder", Dock = DockStyle.Bottom, Height = 24 };
            btnOpenFolder.Click += (s, e) =>
            {
                try { _store.OpenFolderInExplorer(); } catch { }
            };

            btnExportExcel = new Button { Text = "Export Assembly Report (Excel)", Dock = DockStyle.Bottom, Height = 24 };
            btnExportExcel.Click += (s, e) =>
            {
                try
                {
                    // Export into project memory folder (or subfolder)
                    string outDir = System.IO.Path.Combine(_store.RootFolder, "Reports");
                    string path = SwAssemblyExcelReport.ExportActiveDocToExcel(_swApp, outDir, Log);

                    Log($"Excel exported: {path}\r\n");

                    // Open the file for the user
                    try
                    {
                        Process.Start(new ProcessStartInfo
                        {
                            FileName = path,
                            UseShellExecute = true
                        });
                    }
                    catch { }
                }
                catch (Exception ex)
                {
                    Log("Export failed: " + ex.Message + "\r\n");
                }
            };

            btnImportExcel = new Button { Text = "Import Excel & Apply Properties", Dock = DockStyle.Bottom, Height = 24 };
            btnImportExcel.Click += (s, e) =>
            {
                try
                {
                    using (var ofd = new OpenFileDialog())
                    {
                        ofd.Filter = "Excel (*.xlsx)|*.xlsx";
                        ofd.Title = "Select exported report to import";
                        if (ofd.ShowDialog() != DialogResult.OK) return;

                        SwAssemblyExcelReport.ImportExcelAndApplyProperties(_swApp, ofd.FileName, Log);
                        Log("Applied properties from Excel.\r\n");
                    }
                }
                catch (Exception ex)
                {
                    Log("Import failed: " + ex.Message + "\r\n");
                }
            };


            btnClearMemory = new Button { Text = "Clear memory for this project", Dock = DockStyle.Bottom, Height = 24 };
            btnClearMemory.Click += async (s, e) =>
            {
                try
                {
                    await _store.ClearAsync();
                    Log("Memory cleared.\r\n");
                }
                catch (Exception ex)
                {
                    Log("Clear failed: " + ex.Message + "\r\n");
                }
            };

            lblFolder = new Label { Dock = DockStyle.Bottom, Height = 34, Text = "Memory folder: (loading...)" };

            // -------------------------
            // RAG: always-visible toggle
            // -------------------------
            chkRagOnly = new CheckBox
            {
                Text = "Enable RAG mode (use only docs from OllamaRAG folder)",
                Dock = DockStyle.Bottom,
                Checked = false
            };
            chkRagOnly.CheckedChanged += async (s, e) =>
            {
                if (_settings == null) return;

                _settings.EnableRagOnly = chkRagOnly.Checked;
                await _store.SaveSettingsAsync(_settings);

                EnsureRagFolders();
                UpdateRagStatus();
                UpdateRagUiVisibility();
            };

            // -------------------------
            // RAG panel (hidden unless enabled)
            // -------------------------
            pnlRag = new FlowLayoutPanel
            {
                Dock = DockStyle.Bottom,
                FlowDirection = FlowDirection.TopDown,
                WrapContents = false,
                AutoSize = true,
                AutoSizeMode = AutoSizeMode.GrowAndShrink,
                Visible = false,
                Padding = new Padding(6, 6, 6, 6)
            };

            btnSetupGlobalRag = new Button
            {
                Text = "Setup Global RAG Environment (one-time)",
                Width = 320,
                Height = 28
            };
            btnSetupGlobalRag.Click += async (s, e) => await SetupGlobalRagAsync();

            btnOpenRagFolder = new Button
            {
                Text = "Open OllamaRAG folder (project)",
                Width = 320,
                Height = 24
            };
            btnOpenRagFolder.Click += (s, e) =>
            {
                try
                {
                    EnsureRagFolders();
                    string ragFolder = Path.Combine(_store.RootFolder, _settings.RagFolderName ?? "OllamaRAG");
                    if (Directory.Exists(ragFolder))
                    {
                        Process.Start(new ProcessStartInfo
                        {
                            FileName = "explorer.exe",
                            Arguments = $"\"{ragFolder}\"",
                            UseShellExecute = true
                        });
                    }
                }
                catch { }
            };

            btnBuildRagIndex = new Button
            {
                Text = "Build / Refresh RAG index",
                Width = 320,
                Height = 24
            };
            btnBuildRagIndex.Click += async (s, e) => await BuildRagIndexAsync();

            lblRagStatus = new Label
            {
                AutoSize = true,
                Text = "RAG: (loading...)"
            };

            lblOcr = new Label { AutoSize = true, Text = "RAG OCR model (vision):" };
            cboOcrModel = new ComboBox { DropDownStyle = ComboBoxStyle.DropDownList, Width = 320 };
            cboOcrModel.SelectedIndexChanged += async (s, e) =>
            {
                if (_settings == null) return;
                if (cboOcrModel.SelectedItem is string m && !string.IsNullOrWhiteSpace(m))
                {
                    _settings.OcrModelName = m;
                    await _store.SaveSettingsAsync(_settings);
                }
            };

            lblQa = new Label { AutoSize = true, Text = "RAG QA model (answer):" };
            cboQaModel = new ComboBox { DropDownStyle = ComboBoxStyle.DropDownList, Width = 320 };
            cboQaModel.SelectedIndexChanged += async (s, e) =>
            {
                if (_settings == null) return;
                if (cboQaModel.SelectedItem is string m && !string.IsNullOrWhiteSpace(m))
                {
                    _settings.QaModelName = m;
                    await _store.SaveSettingsAsync(_settings);
                }
            };

            lblEmbed = new Label { AutoSize = true, Text = "RAG embedding model:" };
            cboEmbedModel = new ComboBox { DropDownStyle = ComboBoxStyle.DropDownList, Width = 320 };
            cboEmbedModel.SelectedIndexChanged += async (s, e) =>
            {
                if (_settings == null) return;
                if (cboEmbedModel.SelectedItem is string m && !string.IsNullOrWhiteSpace(m))
                {
                    _settings.EmbedModelName = m;
                    await _store.SaveSettingsAsync(_settings);
                }
            };

            pnlRag.Controls.Add(btnSetupGlobalRag);
            pnlRag.Controls.Add(btnOpenRagFolder);
            pnlRag.Controls.Add(btnBuildRagIndex);
            pnlRag.Controls.Add(lblRagStatus);

            pnlRag.Controls.Add(lblOcr);
            pnlRag.Controls.Add(cboOcrModel);
            pnlRag.Controls.Add(lblQa);
            pnlRag.Controls.Add(cboQaModel);
            pnlRag.Controls.Add(lblEmbed);
            pnlRag.Controls.Add(cboEmbedModel);

            // -------------------------
            // Control order
            // -------------------------
            Controls.Add(txtChat);
            Controls.Add(lblStatus);
            Controls.Add(btnSend);
            Controls.Add(txtInput);

            Controls.Add(lblFolder);
            Controls.Add(btnImportExcel);
            Controls.Add(btnExportExcel);
            Controls.Add(btnClearMemory);
            Controls.Add(btnOpenFolder);

            // RAG: toggle always visible, panel hidden until toggle checked
            Controls.Add(pnlRag);
            Controls.Add(chkRagOnly);

            Controls.Add(chkSaveScreenshots);
            Controls.Add(chkIncludeScreenshotInPrompt);
            Controls.Add(chkEnableMemory);

            Controls.Add(btnRefreshModels);
            Controls.Add(cboModels);
        }

        private void UpdateRagUiVisibility()
        {
            bool on = chkRagOnly.Checked;
            pnlRag.Visible = on;
        }

        private void ApplyRagModelSelections()
        {
            if (_settings == null) return;

            if (!string.IsNullOrWhiteSpace(_settings.OcrModelName) && cboOcrModel.Items.Contains(_settings.OcrModelName))
                cboOcrModel.SelectedItem = _settings.OcrModelName;

            if (!string.IsNullOrWhiteSpace(_settings.QaModelName) && cboQaModel.Items.Contains(_settings.QaModelName))
                cboQaModel.SelectedItem = _settings.QaModelName;

            if (!string.IsNullOrWhiteSpace(_settings.EmbedModelName) && cboEmbedModel.Items.Contains(_settings.EmbedModelName))
                cboEmbedModel.SelectedItem = _settings.EmbedModelName;
        }

        private async Task LoadModelsAsync(string selectModel = null)
        {
            btnRefreshModels.Enabled = false;
            cboModels.Enabled = false;
            lblStatus.Text = "Loading models...";

            try
            {
                var models = await _ollama.GetAvailableModelsAsync();

                cboModels.BeginUpdate();
                cboModels.Items.Clear();
                foreach (var m in models) cboModels.Items.Add(m);
                cboModels.EndUpdate();

                // Fill RAG combos too (simple list; user chooses the right ones)
                cboOcrModel.BeginUpdate();
                cboOcrModel.Items.Clear();
                foreach (var m in models) cboOcrModel.Items.Add(m);
                cboOcrModel.EndUpdate();

                cboQaModel.BeginUpdate();
                cboQaModel.Items.Clear();
                foreach (var m in models) cboQaModel.Items.Add(m);
                cboQaModel.EndUpdate();

                cboEmbedModel.BeginUpdate();
                cboEmbedModel.Items.Clear();
                foreach (var m in models) cboEmbedModel.Items.Add(m);
                cboEmbedModel.EndUpdate();

                if (models.Count == 0)
                {
                    lblStatus.Text = "No models found (is Ollama running?)";
                    return;
                }

                if (!string.IsNullOrWhiteSpace(selectModel) && cboModels.Items.Contains(selectModel))
                    cboModels.SelectedItem = selectModel;
                else
                    cboModels.SelectedIndex = 0;

                cboModels.Enabled = true;
                lblStatus.Text = $"Model: {_ollama.Model}";

                // Set defaults if empty
                if (_settings != null)
                {
                    if (string.IsNullOrWhiteSpace(_settings.OcrModelName))
                        _settings.OcrModelName = "qwen3-vl:8b-instruct-q4_K_M";
                    if (string.IsNullOrWhiteSpace(_settings.QaModelName))
                        _settings.QaModelName = "gemma3:12b-it-q4_K_M";
                    if (string.IsNullOrWhiteSpace(_settings.EmbedModelName))
                        _settings.EmbedModelName = "qwen3-embedding:8b-q4_K_M";

                    await _store.SaveSettingsAsync(_settings);
                    ApplyRagModelSelections();
                }
            }
            catch (Exception ex)
            {
                Log($"ERROR loading models: {ex.GetType().Name}: {ex.Message}\r\n");
                lblStatus.Text = "Failed to load models";
            }
            finally
            {
                btnRefreshModels.Enabled = true;
                cboModels.Enabled = cboModels.Items.Count > 0;
            }
        }

        /// <summary>Sends user input to either RAG-only pipeline or normal Ollama chat (optionally with screenshot), then logs to memory.</summary>
        private async Task SendAsync()
        {
            string userText = (txtInput.Text ?? "").Trim();
            if (userText.Length == 0) return;

            txtInput.Clear();
            Log($"You: {userText}\r\n");

            btnSend.Enabled = false;

            try
            {
                await _store.EnsureForActiveDocAsync(_swApp);
                if (_settings == null)
                    _settings = await _store.LoadSettingsAsync();

                lblFolder.Text = $"Memory folder: {_store.RootFolder}";

                // -------------------------
                // RAG-only mode
                // -------------------------
                if (_settings.EnableRagOnly)
                {
                    EnsureRagFolders();
                    UpdateRagStatus();

                    if (!RagRunner.IsIndexReady(_store.RootFolder, _settings))
                    {
                        Log("RAG index not found. Click 'Build / Refresh RAG index' first.\r\n\r\n");
                        return;
                    }

                    lblStatus.Text = "RAG: querying...";
                    var res = await RagRunner.AskAsync(_store.RootFolder, _settings, userText);

                    string ragAnswer = (res.Stdout ?? "").Trim();
                    if (string.IsNullOrWhiteSpace(ragAnswer))
                        ragAnswer = "(empty RAG response)";

                    Log("RAG:\r\n" + ragAnswer + "\r\n\r\n");

                    if (!res.Ok && !string.IsNullOrWhiteSpace(res.Stderr))
                        Log("RAG ERROR:\r\n" + res.Stderr.Trim() + "\r\n\r\n");

                    // ✅ Save RAG Q&A into memory (tagged)
                    if (_settings.EnableMemory)
                    {
                        string activeDocPath = TryGetActiveDocPath();

                        await _store.AppendTurnAsync(new ChatTurn
                        {
                            Role = "user",
                            Content = "[RAG] " + userText,
                            Model = _settings.QaModelName,
                            IncludedScreenshotInPrompt = false,
                            ScreenshotPath = null,
                            ActiveDocPath = activeDocPath
                        });

                        await _store.AppendTurnAsync(new ChatTurn
                        {
                            Role = "assistant",
                            Content = "[RAG] " + ragAnswer,
                            Model = _settings.QaModelName,
                            IncludedScreenshotInPrompt = false,
                            ScreenshotPath = null,
                            ActiveDocPath = activeDocPath
                        });

                        _turnCounterSinceSummary += 2;

                        if (_turnCounterSinceSummary >= Math.Max(4, _settings.SummarizeEveryNTurns))
                        {
                            _turnCounterSinceSummary = 0;
                            await MaybeUpdateSummaryAsync();
                        }
                    }

                    lblStatus.Text = "Done";
                    return;
                }

                // -------------------------
                // Normal chat
                // -------------------------
                string system = await BuildSystemPromptAsync();

                string imageBase64 = null;
                byte[] pngBytes = null;
                string screenshotPath = null;

                // Capture screenshot ONLY if enabled
                if (_settings.IncludeScreenshotInPrompt)
                {
                    lblStatus.Text = "Capturing screenshot...";
                    pngBytes = ScreenshotHelper.CaptureSolidWorksWindowPngBytes(_swApp);

                    if (pngBytes != null && pngBytes.Length > 0)
                    {
                        imageBase64 = Convert.ToBase64String(pngBytes);

                        if (_settings.SaveScreenshotsToDisk)
                            screenshotPath = _store.SaveScreenshotPng(pngBytes, "user");
                    }
                }

                lblStatus.Text = $"Sending to Ollama ({_ollama.Model})...";
                string reply = await _ollama.ChatAsync(system, userText, imageBase64);
                if (string.IsNullOrWhiteSpace(reply)) reply = "(empty response)";

                reply = SanitizePlainText(reply);

                Log($"Ollama: {reply}\r\n\r\n");
                lblStatus.Text = "Done";

                if (_settings.EnableMemory)
                {
                    string activeDocPath = TryGetActiveDocPath();

                    await _store.AppendTurnAsync(new ChatTurn
                    {
                        Role = "user",
                        Content = userText,
                        Model = _ollama.Model,
                        IncludedScreenshotInPrompt = _settings.IncludeScreenshotInPrompt,
                        ScreenshotPath = screenshotPath,
                        ActiveDocPath = activeDocPath
                    });

                    await _store.AppendTurnAsync(new ChatTurn
                    {
                        Role = "assistant",
                        Content = reply,
                        Model = _ollama.Model,
                        IncludedScreenshotInPrompt = _settings.IncludeScreenshotInPrompt,
                        ScreenshotPath = null,
                        ActiveDocPath = activeDocPath
                    });

                    _turnCounterSinceSummary += 2;

                    if (_turnCounterSinceSummary >= Math.Max(4, _settings.SummarizeEveryNTurns))
                    {
                        _turnCounterSinceSummary = 0;
                        await MaybeUpdateSummaryAsync();
                    }
                }
            }
            catch (Exception ex)
            {
                Log($"ERROR: {ex.GetType().Name}: {ex.Message}\r\n\r\n");
                lblStatus.Text = "Error";
            }
            finally
            {
                btnSend.Enabled = true;
            }
        }

        /// <summary>Builds/refreshes the RAG index for documents inside the project RAG folder.</summary>
        private async Task BuildRagIndexAsync()
        {
            try
            {
                await _store.EnsureForActiveDocAsync(_swApp);
                if (_settings == null)
                    _settings = await _store.LoadSettingsAsync();

                EnsureRagFolders();
                UpdateRagStatus();

                lblStatus.Text = "RAG: building index...";
                btnBuildRagIndex.Enabled = false;

                var res = await RagRunner.BuildIndexAsync(_store.RootFolder, _settings);

                if (!string.IsNullOrWhiteSpace(res.Stdout))
                    Log("RAG BUILD:\r\n" + res.Stdout.Trim() + "\r\n\r\n");

                if (!res.Ok && !string.IsNullOrWhiteSpace(res.Stderr))
                    Log("RAG BUILD ERROR:\r\n" + res.Stderr.Trim() + "\r\n\r\n");

                UpdateRagStatus();
                lblStatus.Text = "Done";
            }
            catch (Exception ex)
            {
                Log("RAG build failed: " + ex.Message + "\r\n\r\n");
                lblStatus.Text = "Error";
            }
            finally
            {
                btnBuildRagIndex.Enabled = true;
            }
        }

        /// <summary>Creates required RAG folders inside the project memory root.</summary>
        private void EnsureRagFolders()
        {
            if (_settings == null) return;

            try
            {
                Directory.CreateDirectory(_store.RootFolder);
                Directory.CreateDirectory(Path.Combine(_store.RootFolder, _settings.RagFolderName ?? "OllamaRAG"));
                Directory.CreateDirectory(Path.Combine(_store.RootFolder, _settings.RagStoreFolderName ?? "rag_store"));
            }
            catch { }
        }

        private void UpdateRagStatus()
        {
            if (_settings == null || string.IsNullOrWhiteSpace(_store.RootFolder))
            {
                lblRagStatus.Text = "RAG: (unknown)";
                return;
            }

            bool ready = RagRunner.IsIndexReady(_store.RootFolder, _settings);

            string globalDir = Path.Combine(
                System.Environment.GetFolderPath(System.Environment.SpecialFolder.LocalApplicationData),
                "OllamaCAD"
            );

            string globalScript = Path.Combine(globalDir, "OllamaRAG_Multi.py");
            string globalPython = Path.Combine(globalDir, "venv", "Scripts", "python.exe");

            string ragFolder = Path.Combine(_store.RootFolder, _settings.RagFolderName ?? "OllamaRAG");

            string scriptState = File.Exists(globalScript) ? "script OK" : "script missing";
            string pyState = File.Exists(globalPython) ? "venv OK" : "venv missing";
            string idxState = ready ? "index ready" : "index missing";

            lblRagStatus.Text = $"RAG: {idxState} | {scriptState} | {pyState}\r\nFolder: {ragFolder}";
        }

        // -------------------------
        // Global RAG Setup
        // -------------------------

        /// <summary> Creates/updates the global Python RAG environment (venv + packages + script copy).</summary>
        private async Task SetupGlobalRagAsync()
        {
            btnSetupGlobalRag.Enabled = false;

            try
            {
                lblStatus.Text = "Setting up global RAG...";
                Log("=== Global RAG setup ===\r\n");

                string globalDir = Path.Combine(
                    System.Environment.GetFolderPath(System.Environment.SpecialFolder.LocalApplicationData),
                    "OllamaCAD"
                );

                Directory.CreateDirectory(globalDir);

                string venvDir = Path.Combine(globalDir, "venv");
                string venvPython = Path.Combine(venvDir, "Scripts", "python.exe");

                Log("Global dir: " + globalDir + "\r\n");

                // Copy/refresh script
                EnsureGlobalRagScript(globalDir);

                // Create venv if missing
                if (!Directory.Exists(venvDir))
                {
                    Log("Creating venv...\r\n");
                    var r = await RunProcessAsync("python", "-m venv venv", globalDir);
                    if (!string.IsNullOrWhiteSpace(r.Stdout)) Log(r.Stdout + "\r\n");
                    if (!r.Ok)
                    {
                        if (!string.IsNullOrWhiteSpace(r.Stderr)) Log("ERROR:\r\n" + r.Stderr + "\r\n");
                        Log("Failed to create venv. Make sure Python is installed and available in PATH.\r\n\r\n");
                        return;
                    }
                }
                else
                {
                    Log("Venv already exists.\r\n");
                }

                if (!File.Exists(venvPython))
                {
                    Log("ERROR: venv python not found: " + venvPython + "\r\n\r\n");
                    return;
                }

                // Upgrade pip
                Log("Upgrading pip...\r\n");
                var pipUp = await RunProcessAsync(venvPython, "-m pip install --upgrade pip", globalDir);
                if (!string.IsNullOrWhiteSpace(pipUp.Stdout)) Log(pipUp.Stdout + "\r\n");
                if (!pipUp.Ok)
                {
                    if (!string.IsNullOrWhiteSpace(pipUp.Stderr)) Log("ERROR:\r\n" + pipUp.Stderr + "\r\n");
                    return;
                }

                // Install packages
                Log("Installing packages...\r\n");
                var install = await RunProcessAsync(
                    venvPython,
                    "-m pip install pdf2image pillow faiss-cpu pandas openpyxl numpy tqdm ollama ",
                    globalDir
                );

                if (!string.IsNullOrWhiteSpace(install.Stdout)) Log(install.Stdout + "\r\n");
                if (!install.Ok)
                {
                    if (!string.IsNullOrWhiteSpace(install.Stderr)) Log("ERROR:\r\n" + install.Stderr + "\r\n");
                    return;
                }

                Log("Global RAG environment ready.\r\n\r\n");
                lblStatus.Text = "Global RAG ready";

                // Save python path once (optional)
                try
                {
                    await _store.EnsureForActiveDocAsync(_swApp);
                    if (_settings == null) _settings = await _store.LoadSettingsAsync();

                    if (string.IsNullOrWhiteSpace(_settings.PythonExePath))
                    {
                        _settings.PythonExePath = venvPython;
                        await _store.SaveSettingsAsync(_settings);
                        Log("Saved PythonExePath into project settings.json\r\n\r\n");
                    }
                }
                catch { }

                UpdateRagStatus();
            }
            catch (Exception ex)
            {
                Log("Global setup failed: " + ex.Message + "\r\n\r\n");
                lblStatus.Text = "Error";
            }
            finally
            {
                btnSetupGlobalRag.Enabled = true;
            }
        }

        // Always refresh global script from add-in folder if possible
        private void EnsureGlobalRagScript(string globalDir)
        {
            string globalScript = Path.Combine(globalDir, "OllamaRAG_Multi.py");
            string baseDir = AppDomain.CurrentDomain.BaseDirectory;
            string localScript = Path.Combine(baseDir, "OllamaRAG_Multi.py");

            if (File.Exists(localScript))
            {
                File.Copy(localScript, globalScript, overwrite: true);
                Log("Copied/Updated OllamaRAG_Multi.py in global folder.\r\n");
                return;
            }

            if (File.Exists(globalScript))
            {
                Log("RAG script already in global folder (no local copy to update from).\r\n");
                return;
            }

            Log("ERROR: OllamaRAG_Multi.py not found to copy.\r\n");
            Log("Put OllamaRAG_Multi.py next to your add-in DLL here:\r\n");
            Log(baseDir + "\r\n");
            Log("Then click 'Setup Global RAG Environment' again.\r\n\r\n");
        }

        /// <summary>Builds the system prompt, optionally injecting project memory and selected component properties.</summary>
        private async Task<string> BuildSystemPromptAsync()
        {
            var sb = new StringBuilder();

            sb.AppendLine("You are a helpful SOLIDWORKS assistant.");
            sb.AppendLine("Be practical and concise.");
            sb.AppendLine("OUTPUT RULES:");
            sb.AppendLine("- Plain text only.");
            sb.AppendLine("- Do NOT use markdown formatting (no **, *, #, bullets, code fences).");
            sb.AppendLine("- If you provide code, output it as plain text (no ``` fences).");
            sb.AppendLine("IMAGE RULE:");
            sb.AppendLine("- If an image is provided, analyze what is visible in the screenshot FIRST.");
            //sb.AppendLine("- Then use SOLIDWORKS CONTEXT and the user's question to answer.");
            sb.AppendLine();
            sb.AppendLine("MODEL:");
            sb.AppendLine($"- Current Ollama model: {_ollama.Model}");
            sb.AppendLine();
            sb.AppendLine("SPECIAL OUTPUT MODES:");
            sb.AppendLine("- If the user asks for a SOLIDWORKS VBA macro, output ONLY VBA code (no explanation).");
            sb.AppendLine();


            if (_settings.EnableMemory)
            {
                string summary = await _store.LoadSummaryAsync();
                if (!string.IsNullOrWhiteSpace(summary))
                {
                    sb.AppendLine("PROJECT MEMORY SUMMARY:");
                    sb.AppendLine(summary.Trim());
                    sb.AppendLine();
                }

                var recent = await _store.LoadRecentTurnsAsync(Math.Max(2, _settings.MaxRecentTurns));
                if (recent.Count > 0)
                {
                    sb.AppendLine("RECENT CHAT:");
                    foreach (var t in recent)
                        sb.AppendLine($"{t.Role.ToUpperInvariant()}: {t.Content}");
                    sb.AppendLine();
                }
            }

            //var ctx = SwContextBuilder.Build(_swApp);
            //sb.AppendLine("SOLIDWORKS CONTEXT:");
            //sb.AppendLine($"- Title: {ctx.Title}");
            //sb.AppendLine($"- DocType: {ctx.DocType}");
            //sb.AppendLine($"- SelectionCount: {ctx.SelectionCount}");

            // NEW: inject properties for selected components/parts so the user can ask for
            // mass/material/CoG/etc via prompt without exporting Excel.
            try
            {
                string selectedBlock = SwSelectionProperties.BuildSelectedPartsContext(_swApp);
                if (!string.IsNullOrWhiteSpace(selectedBlock))
                {
                    sb.AppendLine();
                    sb.AppendLine("SELECTED COMPONENT PROPERTIES:");
                    sb.AppendLine(selectedBlock);
                }
            }
            catch
            {
                // ignore selection context errors
            }

            return sb.ToString();
        }

        /// <summary>Updates the rolling memory summary using recent chat turns.</summary>
        private async Task MaybeUpdateSummaryAsync()
        {
            try
            {
                lblStatus.Text = "Updating memory summary...";

                string existing = await _store.LoadSummaryAsync();
                var recent = await _store.LoadRecentTurnsAsync(20);

                string updated = await _summarizer.UpdateSummaryAsync(existing, recent.ToArray());
                await IoCompat.WriteAllTextAsync(_store.SummaryPath, updated ?? "", Encoding.UTF8);

                lblStatus.Text = "Done";
            }
            catch (Exception ex)
            {
                Log("Summary update failed: " + ex.Message + "\r\n");
                lblStatus.Text = "Done";
            }
        }

        private string TryGetActiveDocPath()
        {
            try
            {
                var doc = (ModelDoc2)_swApp.ActiveDoc;
                if (doc == null) return null;
                return doc.GetPathName();
            }
            catch { return null; }
        }

        private static string SanitizePlainText(string s)
        {
            if (string.IsNullOrWhiteSpace(s)) return s;
            s = s.Replace("```", "");
            return s.Trim();
        }

        private void Log(string s) => txtChat.AppendText(s);

        // Process helper (.NET Framework 4.8 friendly)
        private sealed class ProcResult
        {
            public int ExitCode { get; set; }
            public string Stdout { get; set; }
            public string Stderr { get; set; }
            public bool Ok => ExitCode == 0;
        }

        private async Task<ProcResult> RunProcessAsync(string exe, string args, string workDir)
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
    }
}
