using System;
using System.Net.Http;
using System.Text;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System.Threading.Tasks;
using System.Collections.Generic;
using System.Linq;

namespace OllamaCAD
{   
    /// <summary>
    /// Lightweight HTTP client wrapper for communicating with a local Ollama server.
    ///
    /// Responsibilities:
    /// - Queries installed models from Ollama (GET /api/tags).
    /// - Sends chat requests to Ollama (POST /api/chat).
    /// - Supports optional image input (base64) for vision-capable models.
    /// - Returns plain text responses for use inside the SOLIDWORKS add-in.
    ///
    /// Designed for local-only LLM/VLM usage (no cloud dependency).
    /// </summary>
    internal sealed class OllamaClient : IDisposable
    {
        private readonly HttpClient _http;

        // Put your installed model here by default:
        public string Model { get; set; } = "qwen3-vl:8b-instruct-q4_K_M";

        public OllamaClient(string baseUrl)
        {
            _http = new HttpClient
            {
                BaseAddress = new Uri(baseUrl),
                Timeout = TimeSpan.FromSeconds(120)
            };
        }

        // NEW: returns installed models from Ollama (GET /api/tags)
        public async Task<List<string>> GetAvailableModelsAsync()
        {
            HttpResponseMessage resp = await _http.GetAsync("/api/tags");
            resp.EnsureSuccessStatusCode();

            string body = await resp.Content.ReadAsStringAsync();
            JObject parsed = JObject.Parse(body);

            var result = new List<string>();

            JToken modelsToken = parsed["models"];
            if (modelsToken != null && modelsToken.Type == JTokenType.Array)
            {
                foreach (JToken m in modelsToken)
                {
                    JToken nameToken = m["name"];
                    if (nameToken != null)
                    {
                        string name = nameToken.ToString();
                        if (!string.IsNullOrWhiteSpace(name) && !result.Contains(name))
                            result.Add(name);
                    }
                }
            }

            result.Sort();
            return result;
        }

        public async Task<string> ChatAsync(string system, string user, string imageBase64 = null)
        {
            JArray messages = new JArray
            {
                new JObject
                {
                    ["role"] = "system",
                    ["content"] = system
                }
            };

            JObject userMsg = new JObject
            {
                ["role"] = "user",
                ["content"] = user
            };

            if (!string.IsNullOrWhiteSpace(imageBase64))
            {
                userMsg["images"] = new JArray(imageBase64);
            }
            messages.Add(userMsg);

            JObject payload = new JObject
            {
                ["model"] = Model,
                ["stream"] = false,
                ["messages"] = messages
            };

            string json = payload.ToString(Formatting.None);

            // Absolute path avoids BaseAddress + relative path surprises
            HttpResponseMessage resp = await _http.PostAsync(
                "/api/chat",
                new StringContent(json, Encoding.UTF8, "application/json")
            );

            resp.EnsureSuccessStatusCode();

            string body = await resp.Content.ReadAsStringAsync();
            JObject parsed = JObject.Parse(body);

            return parsed.SelectToken("message.content")?.ToString() ?? "";
        }

        public void Dispose()
        {
            _http?.Dispose();
        }
    }
}
