// Axon Outlook add-in — Model/provider config + chat calls (LoadConfig, ModelComplete, CallChat, scrub).  (partial of Connect; split out of AxonAddin.cs.)
using System;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using Microsoft.Win32;
using System.Windows.Forms;
using Extensibility;
using Microsoft.Office.Core;

namespace Axon.OutlookAddin
{
    public partial class Connect
    {
        private static string _apiBase, _apiKey, _model;
        private static string _backupBase, _backupKey, _backupModel;

        // Config from %APPDATA%\AxonOutlook\config.json:
        //   { api_base, api_key, model, backup_api_base?, backup_api_key?, backup_model? }
        // The primary provider (Mistral in the bundled deployment) is tried first; if it errors or is
        // down, the BACKUP provider (OpenAI by default) transparently takes over so the add-in keeps
        // working. If no api_key/backup_api_key is set, we fall back to the OpenAI key baked into a
        // co-located Axon app (.env) — that's the bundled-with-the-dot case. On-site deployments point
        // api_base at their own model server.
        private void LoadConfig()
        {
            _apiBase = "https://api.openai.com/v1";
            _apiKey = "";
            _model = "gpt-4o";   // better at matching email topic -> folder than mini (config can override)
            // Backup defaults to OpenAI cloud; used only when the primary call fails.
            _backupBase = "https://api.openai.com/v1";
            _backupKey = "";
            _backupModel = "gpt-4o";
            try
            {
                string p = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
                                        "AxonOutlook", "config.json");
                if (File.Exists(p))
                {
                    var js = new System.Web.Script.Serialization.JavaScriptSerializer();
                    var d = (System.Collections.Generic.Dictionary<string, object>)js.DeserializeObject(File.ReadAllText(p));
                    if (d.ContainsKey("api_base") && d["api_base"] != null) _apiBase = d["api_base"].ToString().TrimEnd('/');
                    if (d.ContainsKey("api_key") && d["api_key"] != null) _apiKey = d["api_key"].ToString();
                    if (d.ContainsKey("model") && d["model"] != null) _model = d["model"].ToString();
                    if (d.ContainsKey("backup_api_base") && d["backup_api_base"] != null) _backupBase = d["backup_api_base"].ToString().TrimEnd('/');
                    if (d.ContainsKey("backup_api_key") && d["backup_api_key"] != null) _backupKey = d["backup_api_key"].ToString();
                    if (d.ContainsKey("backup_model") && d["backup_model"] != null) _backupModel = d["backup_model"].ToString();
                }
            }
            catch { }
            string baked = BakedKey();   // OpenAI key from a co-located Axon .env (bundled deployment)
            if (string.IsNullOrEmpty(_apiKey)) _apiKey = baked;
            if (string.IsNullOrEmpty(_backupKey)) _backupKey = baked;   // OpenAI is the default backup
        }

        // Read OPENAI_API_KEY from a co-located Axon app's .env (bundled-with-the-dot deployment).
        private string BakedKey()
        {
            try
            {
                string dir = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
                string root = Path.GetFullPath(Path.Combine(dir, ".."));
                string[] cands = { Path.Combine(root, "_internal", ".env"), Path.Combine(root, ".env"),
                                   Path.Combine(root, "assistant", ".env") };
                foreach (var p in cands)
                    if (File.Exists(p))
                        foreach (var line in File.ReadAllLines(p))
                            if (line.TrimStart().StartsWith("OPENAI_API_KEY"))
                            {
                                int i = line.IndexOf('=');
                                if (i > 0) { string k = line.Substring(i + 1).Trim().Trim('"'); if (k.Length > 10) return k; }
                            }
            }
            catch { }
            return "";
        }

        // One chat completion -> the assistant's text (or null). Tries the primary provider (Mistral in
        // the bundled deployment); if it errors, times out, or is down, it transparently retries on the
        // backup provider (OpenAI) so the feature keeps working. Shared by suggest, summarize, and reply.
        private string ModelComplete(string prompt, double temperature)
        {
            LoadConfig();
            string text = CallChat(_apiBase, _apiKey, _model, prompt, temperature);
            if (string.IsNullOrEmpty(text))
            {
                // Primary failed / unreachable / empty -> fall over to the backup provider, but only if
                // it is usable and actually a different endpoint (no point re-calling the same one).
                bool backupUsable = !string.IsNullOrEmpty(_backupKey);
                bool backupDifferent = !(string.Equals(_backupBase, _apiBase, StringComparison.OrdinalIgnoreCase)
                                         && string.Equals(_backupModel, _model, StringComparison.OrdinalIgnoreCase));
                if (backupUsable && backupDifferent)
                    text = CallChat(_backupBase, _backupKey, _backupModel, prompt, temperature);
            }
            return ScrubIdentity(text);
        }

        // Axon must never reveal the underlying model/provider in any output. Strip brand names the
        // model might emit (OpenAI, ChatGPT, GPT-4o, Mistral, ...) from summaries, replies, and drafts.
        private static string ScrubIdentity(string s)
        {
            if (string.IsNullOrEmpty(s)) return s;
            var ic = System.Text.RegularExpressions.RegexOptions.IgnoreCase;
            s = System.Text.RegularExpressions.Regex.Replace(s, @"\bchat\s?gpt\b", "Axon", ic);
            s = System.Text.RegularExpressions.Regex.Replace(s, @"\bopen\s?ai\b", "Axon", ic);
            s = System.Text.RegularExpressions.Regex.Replace(s, @"\bmistral(\s?ai)?\b", "Axon", ic);
            s = System.Text.RegularExpressions.Regex.Replace(s, @"\bgpt[-\s]?[\w.]+", "Axon", ic);
            return s;
        }

        // One OpenAI-compatible chat call against a specific provider -> the assistant's text (or null
        // on any error/non-success). Kept provider-agnostic so ModelComplete can call primary then backup.
        private string CallChat(string apiBase, string apiKey, string model, string prompt, double temperature)
        {
            try
            {
                var js = new System.Web.Script.Serialization.JavaScriptSerializer();
                var reqObj = new System.Collections.Generic.Dictionary<string, object>
                {
                    { "model", model },
                    { "temperature", temperature },
                    { "stream", false },
                    { "messages", new object[] { new System.Collections.Generic.Dictionary<string, object> { { "role", "user" }, { "content", prompt } } } }
                };
                System.Net.ServicePointManager.SecurityProtocol |= System.Net.SecurityProtocolType.Tls12;
                using (var http = new System.Net.Http.HttpClient())
                {
                    http.Timeout = TimeSpan.FromSeconds(60);   // fail over to the backup reasonably fast
                    if (!string.IsNullOrEmpty(apiKey))
                        http.DefaultRequestHeaders.TryAddWithoutValidation("Authorization", "Bearer " + apiKey);
                    var content = new System.Net.Http.StringContent(js.Serialize(reqObj), System.Text.Encoding.UTF8, "application/json");
                    var resp = http.PostAsync(apiBase + "/chat/completions", content).Result;
                    if (!resp.IsSuccessStatusCode) return null;
                    var d = (System.Collections.Generic.Dictionary<string, object>)js.DeserializeObject(resp.Content.ReadAsStringAsync().Result);
                    var choices = d.ContainsKey("choices") ? d["choices"] as object[] : null;
                    if (choices == null || choices.Length == 0) return null;
                    var msg = (System.Collections.Generic.Dictionary<string, object>)((System.Collections.Generic.Dictionary<string, object>)choices[0])["message"];
                    return (msg != null && msg.ContainsKey("content") && msg["content"] != null) ? msg["content"].ToString() : null;
                }
            }
            catch { return null; }
        }
    }
}
