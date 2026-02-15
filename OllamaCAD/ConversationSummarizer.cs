using SolidWorks.Interop.sldworks;
using System;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OllamaCAD
{   
    /// <summary>
    /// Generates and updates a compact “project memory” summary for the SOLIDWORKS CAD assistant.
    ///
    /// - Takes the existing summary plus a set of recent chat turns.
    /// - Prompts the local Ollama model to produce an updated, concise summary (plain text).
    /// - Focuses on CAD-relevant context: user intent, constraints, decisions, paths, chosen models/tools, and TODOs.
    /// - Uses text-only summarization (no images).
    /// </summary>
    internal sealed class ConversationSummarizer
    {
        private readonly OllamaClient _ollama;

        public ConversationSummarizer(OllamaClient ollama)
        {
            _ollama = ollama;
        }

        public async Task<string> UpdateSummaryAsync(string existingSummary, ChatTurn[] recentTurns)
        {
            // Keep it small and “CAD assistant relevant”
            string turnsText = string.Join("\n",
                recentTurns.Select(t => $"{t.Role.ToUpperInvariant()}: {t.Content}")
            );

            string system =
                "You are a memory summarizer for a CAD assistant.\n" +
                "Goal: produce a compact summary to help future responses.\n" +
                "Rules:\n" +
                "- Output plain text only.\n" +
                "- Keep under ~1200 words.\n" +
                "- Preserve: user intent, constraints, decisions, file/folder paths, model/tool choices, and open TODOs.\n" +
                "- Do NOT include irrelevant chat.\n";

            string user =
                "EXISTING SUMMARY (may be empty):\n" +
                existingSummary + "\n\n" +
                "NEWEST TURNS:\n" + turnsText + "\n\n" +
                "Write an UPDATED SUMMARY:";

            // text-only summarization (no images)
            return await _ollama.ChatAsync(system, user, null);
        }
    }
}
