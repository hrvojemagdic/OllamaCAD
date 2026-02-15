OllamaCAD - Installation Guide (SOLIDWORKS Add-in)
=================================================

This add-in integrates local Ollama LLM/VLM models inside SOLIDWORKS via a Task Pane UI.
No Visual Studio is required for installation.

REQUIREMENTS
------------
1) SOLIDWORKS 2020+ (64-bit)
2) .NET Framework 4.7.2+ (usually already installed on Windows 10/11)
3) Ollama running locally:
   - Start Ollama
   - Default endpoint used by the add-in: http://localhost:11434/

FILES IN THIS FOLDER
--------------------
- OllamaCAD.dll
- OllamaCAD_Addin.reg
- install.bat
- uninstall.bat

INSTALL (RECOMMENDED)
--------------------
1) Close SOLIDWORKS
2) Right-click install.bat -> "Run as administrator"
3) Start SOLIDWORKS
4) Go to: Tools -> Add-Ins...
5) Enable "Ollama CAD" (Active Add-ins)
   Optionally enable at Startup.

If you do not see the add-in in Tools -> Add-Ins:
- Ensure install.bat was run as Administrator
- Ensure you used the 64-bit regasm (the script selects Framework64 automatically)

UNINSTALL
---------
1) Close SOLIDWORKS
2) Right-click uninstall.bat -> "Run as administrator"
3) Start SOLIDWORKS (the add-in should no longer be listed)

TROUBLESHOOTING
---------------
- "No models found (is Ollama running?)"
  -> Start Ollama and ensure http://localhost:11434/ is reachable.

- Task pane is empty / not showing
  -> Try enabling the add-in again via Tools -> Add-Ins.
  -> Ensure OllamaCAD.png is in the same folder as OllamaCAD.dll if your build expects it.

- COM registration issues
  -> Run install.bat as Administrator.
  -> Confirm that %WINDIR%\Microsoft.NET\Framework64\v4.0.30319\regasm.exe exists.

RAG MODE (OPTIONAL)
-------------------
RAG mode requires Python and additional packages. In the add-in UI:
- Click "Setup Global RAG Environment (one-time)"
- Put documents into the project "OllamaRAG" folder
- Click "Build / Refresh RAG index"
- Enable "RAG mode" checkbox

SECURITY / PRIVACY
------------------
OllamaCAD is designed for local-only workflows:
- No cloud calls by default
- No telemetry
- Project memory is stored in per-document folders on disk
