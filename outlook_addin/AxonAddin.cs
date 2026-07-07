// Axon Outlook add-in — adds two ribbon buttons ("File with Axon", "Download with Axon").
// Implemented as a managed COM add-in against the REAL Office interop interfaces
// (Extensibility.IDTExtensibility2, Microsoft.Office.Core.IRibbonExtensibility). Those types are
// EMBEDDED at build time (csc /link), so the compiled DLL is self-contained and needs no PIAs on
// the user's machine. Outlook objects are used late-bound (dynamic), so there's no hard dependency
// on the Outlook object model either.
//
// Phase 1: the buttons appear and report the selected email's subject (proves load + email access).

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
    // Converts a managed Image into the COM IPictureDisp the ribbon's getImage callback needs.
    internal class RibbonImage : AxHost
    {
        private RibbonImage() : base("00000000-0000-0000-0000-000000000000") { }
        public static stdole.IPictureDisp Get(System.Drawing.Image img)
        {
            return (stdole.IPictureDisp)AxHost.GetIPictureDispFromPicture(img);
        }
    }

    [ComVisible(true)]
    [Guid("7B2C9E14-6A3D-4F58-9C21-3E5A1B7D4F60")]
    [ProgId("Axon.OutlookAddin")]
    [ClassInterface(ClassInterfaceType.AutoDispatch)]   // exposes the ribbon callbacks via IDispatch
    public class Connect : IDTExtensibility2, IRibbonExtensibility
    {
        private object _app;   // Outlook.Application (late-bound)

        // --- IDTExtensibility2 ---
        public void OnConnection(object Application, ext_ConnectMode ConnectMode, object AddInInst, ref Array custom) { _app = Application; }
        public void OnDisconnection(ext_DisconnectMode RemoveMode, ref Array custom) { _app = null; }
        public void OnAddInsUpdate(ref Array custom) { }
        public void OnStartupComplete(ref Array custom) { }
        public void OnBeginShutdown(ref Array custom) { }

        // --- IRibbonExtensibility ---
        public string GetCustomUI(string RibbonID)
        {
            // Axon lives ONLY in the right-click menu (no ribbon buttons). Add it to the menu you get
            // on an email in the list, on multiple selected emails, and inside an open/previewed email.
            if (RibbonID == "Microsoft.Outlook.Explorer")
                return CtxUI(CtxMenu("ContextMenuMailItem") + CtxMenu("ContextMenuReadOnlyMailText"));
            if (RibbonID == "Microsoft.Outlook.Mail.Read")
                return CtxUI(CtxMenu("ContextMenuReadOnlyMailText"));
            if (RibbonID == "Microsoft.Outlook.Mail.Compose")
                return ComposeRibbon();   // compose body right-click is Word's, so use a ribbon button
            return null;
        }

        private string CtxUI(string menus)
        {
            return "<customUI xmlns='http://schemas.microsoft.com/office/2009/07/customui'>" +
                   "<contextMenus>" + menus + "</contextMenus></customUI>";
        }

        // A small "Send Later" button in an Axon group on the compose Message tab (reliable — the
        // compose body's right-click menu belongs to the Word editor and can't be extended).
        private string ComposeRibbon()
        {
            return "<customUI xmlns='http://schemas.microsoft.com/office/2009/07/customui'>" +
                   "<ribbon><tabs><tab idMso='TabNewMailMessage'>" +
                   "<group id='axonComposeGroup' label='Axon'>" +
                   "<button id='axonWriteBtn' label='Write with Axon' size='large' " +
                   "getImage='GetWriteImage' onAction='OnWriteEmail'/>" +
                   "<button id='axonSendLaterBtn' label='Send Later' size='large' " +
                   "getImage='GetSendLaterImage' onAction='OnSendLater'/>" +
                   "</group></tab></tabs></ribbon></customUI>";
        }

        // Axon's right-click items for a given Office context-menu id (button ids must be unique).
        // Summarize/Reply have no icon (avoids any invalid-image risk); Move/Download keep theirs.
        private string CtxMenu(string idMso)
        {
            return "<contextMenu idMso='" + idMso + "'>" +
                   "<menuSeparator id='axonSep_" + idMso + "'/>" +
                   "<button id='axonSummarize_" + idMso + "' label='Summarize with Axon' getImage='GetSummarizeImage' onAction='OnSummarize'/>" +
                   "<button id='axonReply_" + idMso + "' label='Reply with Axon' getImage='GetReplyImage' onAction='OnReply'/>" +
                   "<button id='axonSchedule_" + idMso + "' label='Schedule with Axon' getImage='GetScheduleImage' onAction='OnSchedule'/>" +
                   "<button id='axonFollowUp_" + idMso + "' label='Follow up with Axon' getImage='GetFollowUpImage' onAction='OnFollowUp'/>" +
                   "<button id='axonAttach_" + idMso + "' label='Forward as attachment' getImage='GetAttachImage' onAction='OnAttachEmail'/>" +
                   "<button id='axonMove_" + idMso + "' label='Move with Axon' getImage='GetMoveImage' onAction='OnFile'/>" +
                   "<button id='axonDownload_" + idMso + "' label='Download with Axon' getImage='GetDownloadImage' onAction='OnDownload'/>" +
                   "</contextMenu>";
        }


        // --- custom ribbon image (the Axon-branded Move icon, distinct from Outlook's built-ins) ---
        private System.Drawing.Image _moveIcon, _downloadIcon, _summarizeIcon, _replyIcon, _scheduleIcon,
                                     _followUpIcon, _sendLaterIcon, _writeIcon, _attachIcon;

        private System.Drawing.Image LoadIcon(string file, ref System.Drawing.Image cache)
        {
            if (cache == null)
            {
                string dir = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
                byte[] bytes = File.ReadAllBytes(Path.Combine(dir, file));   // ReadAllBytes -> no file lock
                cache = System.Drawing.Image.FromStream(new System.IO.MemoryStream(bytes));
            }
            return cache;
        }

        public stdole.IPictureDisp GetMoveImage(object control)
        {
            try { return RibbonImage.Get(LoadIcon("axon-move.png", ref _moveIcon)); } catch { return null; }
        }

        public stdole.IPictureDisp GetDownloadImage(object control)
        {
            try { return RibbonImage.Get(LoadIcon("axon-download.png", ref _downloadIcon)); } catch { return null; }
        }

        public stdole.IPictureDisp GetSummarizeImage(object control)
        {
            try { return RibbonImage.Get(LoadIcon("axon-summarize.png", ref _summarizeIcon)); } catch { return null; }
        }

        public stdole.IPictureDisp GetReplyImage(object control)
        {
            try { return RibbonImage.Get(LoadIcon("axon-reply.png", ref _replyIcon)); } catch { return null; }
        }

        public stdole.IPictureDisp GetScheduleImage(object control)
        {
            try { return RibbonImage.Get(LoadIcon("axon-schedule.png", ref _scheduleIcon)); } catch { return null; }
        }

        public stdole.IPictureDisp GetFollowUpImage(object control)
        {
            try { return RibbonImage.Get(LoadIcon("axon-followup.png", ref _followUpIcon)); } catch { return null; }
        }

        public stdole.IPictureDisp GetSendLaterImage(object control)
        {
            try { return RibbonImage.Get(LoadIcon("axon-sendlater.png", ref _sendLaterIcon)); } catch { return null; }
        }

        public stdole.IPictureDisp GetWriteImage(object control)
        {
            try { return RibbonImage.Get(LoadIcon("axon-write.png", ref _writeIcon)); } catch { return null; }
        }

        public stdole.IPictureDisp GetAttachImage(object control)
        {
            try { return RibbonImage.Get(LoadIcon("axon-attach.png", ref _attachIcon)); } catch { return null; }
        }

        // --- ribbon button callbacks (Office invokes these by name via IDispatch) ---
        public void OnFile(object control)
        {
            try
            {
                object m = GetSelectedMail();
                if (m == null) { Ui.Notify("Select an email first.", "Axon intelligence"); return; }
                dynamic mail = m;
                string[] folders = EnumerateInboxFolders();
                if (folders.Length == 0)
                {
                    Ui.Notify("You have no Inbox subfolders yet. Create some folders to move emails into.", "Axon intelligence");
                    return;
                }
                string subject = ""; try { subject = (string)mail.Subject; } catch { }
                string sender = ""; try { sender = (string)mail.SenderName; } catch { }
                try { string em = (string)mail.SenderEmailAddress; if (!string.IsNullOrEmpty(em)) sender = (sender + " <" + em + ">").Trim(); } catch { }
                string body = ""; try { body = (string)mail.Body; } catch { }
                // Show the picker IMMEDIATELY; fetch AI suggestions on a background thread and fill them in.
                var dlg = new FolderPicker(subject, folders);
                var worker = new System.Threading.Thread(() =>
                {
                    Filing fil = SuggestFiling(subject, sender, body, folders);
                    dlg.SetSuggestions(fil.matches, fil.newFolder);
                });
                worker.IsBackground = true;
                worker.Start();
                try
                {
                    if (dlg.ShowDialog() == DialogResult.OK)
                    {
                        if (!string.IsNullOrEmpty(dlg.CreateFolder))
                        {
                            dynamic dest = CreateInboxSubfolder(dlg.CreateFolder.Trim());
                            if (dest != null) mail.Move(dest);
                        }
                        else if (!string.IsNullOrEmpty(dlg.Chosen))
                        {
                            MoveTo(mail, dlg.Chosen);
                        }
                    }
                }
                finally { dlg.Dispose(); }
            }
            catch (Exception ex) { Ui.Notify("Axon error: " + ex.Message, "Axon intelligence"); }
        }

        // Maps each folder's display path (e.g. "Clients / Acme / Invoices") to its EntryID, so we can
        // move into deeply-nested folders unambiguously (folder names may even contain "/").
        private System.Collections.Generic.Dictionary<string, string> _folderMap =
            new System.Collections.Generic.Dictionary<string, string>();

        // ALL folders under the Inbox, recursively, as display paths (in-process COM, fast).
        private string[] EnumerateInboxFolders()
        {
            _folderMap = new System.Collections.Generic.Dictionary<string, string>();
            var list = new System.Collections.Generic.List<string>();
            try
            {
                dynamic ns = ((dynamic)_app).GetNamespace("MAPI");
                dynamic inbox = ns.GetDefaultFolder(6);   // olFolderInbox
                CollectFolders(inbox, "", list, 0);
            }
            catch { }
            return list.ToArray();
        }

        private void CollectFolders(dynamic parent, string prefix, System.Collections.Generic.List<string> list, int depth)
        {
            if (depth > 8) return;
            dynamic subs;
            try { subs = parent.Folders; } catch { return; }
            foreach (dynamic f in subs)
            {
                string name; try { name = (string)f.Name; } catch { continue; }
                string disp = prefix == "" ? name : prefix + " / " + name;
                try { _folderMap[disp] = (string)f.EntryID; } catch { }
                list.Add(disp);
                CollectFolders(f, disp, list, depth + 1);   // recurse into nested subfolders
            }
        }

        // Move the email into the chosen folder (looked up by EntryID from its display path).
        private void MoveTo(dynamic mail, string display)
        {
            try
            {
                string eid;
                if (_folderMap != null && _folderMap.TryGetValue(display, out eid))
                {
                    dynamic ns = ((dynamic)_app).GetNamespace("MAPI");
                    dynamic dest = ns.GetFolderFromID(eid);
                    if (dest != null) { mail.Move(dest); return; }
                }
                Ui.Notify("Folder not found: " + display, "Axon intelligence");
            }
            catch (Exception ex) { Ui.Notify("Couldn't move: " + ex.Message, "Axon intelligence"); }
        }

        private class Filing { public string[] matches = new string[0]; public string newFolder = ""; }

        // Suggest folders via an OpenAI-COMPATIBLE chat API. `api_base` can point at OpenAI's cloud
        // or any local server that speaks the same protocol (Ollama, vLLM, LM Studio, LocalAI, ...),
        // so the same add-in works for cloud or fully on-site deployments — only the config differs.
        // If the API is unreachable the picker still lists every folder; only AI ranking is skipped.
        private Filing SuggestFiling(string subject, string sender, string body, string[] folders)
        {
            var result = new Filing();
            try { TrySuggestViaApi(subject, sender, body, folders, result); } catch { }
            return result;
        }

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

        private bool TrySuggestViaApi(string subject, string sender, string body, string[] folders, Filing result)
        {
            var js = new System.Web.Script.Serialization.JavaScriptSerializer();
            string b = body ?? "";
            if (b.Length > 1500) b = b.Substring(0, 1500);
            string prompt =
                "You are filing an email into one of the user's Outlook folders.\n\n" +
                "Email\n  Subject: " + subject + "\n  From: " + sender + "\n  Body (truncated):\n" + b + "\n\n" +
                "The user's folders (name, with full path for nested ones):\n" + js.Serialize(folders) + "\n\n" +
                "Decide based on what the email is ABOUT and WHO it is from. Pick the folder whose name/path " +
                "most closely matches that topic; the sender's company/domain is a strong hint. List the " +
                "best-fitting folders first (up to 5), but ONLY include a folder that is a genuinely good fit " +
                "(do not pad the list). If none clearly fit, leave matches empty and propose a short new folder.\n" +
                "Reply with ONLY JSON: {\"matches\": [up to 5 folder names copied EXACTLY from the list, best " +
                "first], \"new_folder\": \"a short (1-3 word) new folder name if nothing fits, else an empty string\"}";
            string text = ModelComplete(prompt, 0);
            if (string.IsNullOrEmpty(text)) return false;
            var mt = System.Text.RegularExpressions.Regex.Match(text, "\\{[\\s\\S]*\\}");
            if (!mt.Success) return false;
            ParseSuggestion(mt.Value, folders, result, js);
            return true;
        }

        // Map the model's JSON onto the real folder names (case-insensitive); if it "invents" a new
        // folder that already exists, treat it as a match instead.
        private void ParseSuggestion(string json, string[] folders, Filing result, System.Web.Script.Serialization.JavaScriptSerializer js)
        {
            var d = (System.Collections.Generic.Dictionary<string, object>)js.DeserializeObject(json);
            var lower = new System.Collections.Generic.Dictionary<string, string>();
            foreach (var f in folders) if (f != null) lower[f.ToLowerInvariant()] = f;
            var matches = new System.Collections.Generic.List<string>();
            if (d.ContainsKey("matches") && d["matches"] is object[])
                foreach (var x in (object[])d["matches"])
                {
                    if (x == null) continue;
                    string k = x.ToString().Trim().ToLowerInvariant();
                    if (lower.ContainsKey(k) && !matches.Contains(lower[k])) matches.Add(lower[k]);
                }
            string nf = (d.ContainsKey("new_folder") && d["new_folder"] != null) ? d["new_folder"].ToString().Trim() : "";
            if (nf.Length > 0)
            {
                string nfl = nf.ToLowerInvariant();
                if (lower.ContainsKey(nfl)) { if (!matches.Contains(lower[nfl])) matches.Add(lower[nfl]); }
                else result.newFolder = nf;
            }
            if (matches.Count > 5) matches = matches.GetRange(0, 5);
            result.matches = matches.ToArray();
        }

        // Create (or reuse) a top-level Inbox subfolder by name; returns the folder or null.
        private dynamic CreateInboxSubfolder(string name)
        {
            try
            {
                dynamic ns = ((dynamic)_app).GetNamespace("MAPI");
                dynamic inbox = ns.GetDefaultFolder(6);
                foreach (dynamic f in inbox.Folders)
                    if (string.Equals((string)f.Name, name, StringComparison.OrdinalIgnoreCase)) return f;
                return inbox.Folders.Add(name);
            }
            catch (Exception ex) { Ui.Notify("Couldn't create folder: " + ex.Message, "Axon intelligence"); return null; }
        }

        private string _archiveBaseDir;
        private string _archiveCompany;   // client/company of the email being archived (for the memory)

        // --- archive memory: remember where each client's emails were last filed -----------------
        private static string ArchiveMemPath()
        {
            return Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
                                "AxonOutlook", "archive_memory.json");
        }

        private System.Collections.Generic.Dictionary<string, object> ReadArchiveMemory()
        {
            try
            {
                string p = ArchiveMemPath();
                if (File.Exists(p))
                {
                    var js = new System.Web.Script.Serialization.JavaScriptSerializer();
                    var d = js.DeserializeObject(File.ReadAllText(p)) as System.Collections.Generic.Dictionary<string, object>;
                    if (d != null) return d;
                }
            }
            catch { }
            return new System.Collections.Generic.Dictionary<string, object>();
        }

        private void SaveArchiveChoice(string company, string relPath)
        {
            try
            {
                if (string.IsNullOrWhiteSpace(company) || string.IsNullOrWhiteSpace(relPath)) return;
                if (Path.IsPathRooted(relPath)) return;   // only remember relative archive paths
                var mem = ReadArchiveMemory();
                mem[company.Trim().ToLowerInvariant()] = relPath;
                var js = new System.Web.Script.Serialization.JavaScriptSerializer();
                Directory.CreateDirectory(Path.GetDirectoryName(ArchiveMemPath()));
                File.WriteAllText(ArchiveMemPath(), js.Serialize(mem));
            }
            catch { }
        }

        private string RememberedFolder(string company)
        {
            try
            {
                if (string.IsNullOrWhiteSpace(company)) return "";
                object v;
                if (ReadArchiveMemory().TryGetValue(company.Trim().ToLowerInvariant(), out v) && v != null)
                    return v.ToString();
            }
            catch { }
            return "";
        }

        public void OnDownload(object control)
        {
            try
            {
                object m = GetSelectedMail();
                if (m == null) { Ui.Notify("Select an email first."); return; }
                dynamic mail = m;
                string subject = ""; try { subject = (string)mail.Subject; } catch { }
                var cfg = ReadArchiveCfg();
                if (!cfg.Ready)
                {
                    // No archive configured -> the simple 'pick a save folder' picker.
                    var folders = LoadDownloadFolders();
                    var dlg0 = new DownloadPicker(subject, folders);
                    try
                    {
                        var r0 = dlg0.ShowDialog();
                        SaveDownloadFolders(dlg0.Folders);
                        if (r0 == DialogResult.OK && !string.IsNullOrEmpty(dlg0.Chosen))
                            SaveEmail(mail, dlg0.Chosen, subject);
                    }
                    finally { dlg0.Dispose(); }
                    return;
                }

                // Archive flow: show the picker immediately, extract entities + suggest on a thread.
                var picker = new FolderPicker(subject, new string[0], "Download email",
                    "Save this email to a folder", "Create && Save", "Save", "Cancel");
                var worker = new System.Threading.Thread(() =>
                {
                    try
                    {
                        var info = ExtractArchiveInfo(mail, cfg);
                        string code = Field(info, "code"), company = Field(info, "company"),
                               type = Field(info, "type"), year = Field(info, "year"), sap = Field(info, "sap");
                        if (string.IsNullOrWhiteSpace(year)) year = DateTime.Now.Year.ToString();
                        string sender = ""; try { sender = (string)mail.SenderName; } catch { }
                        string body = ""; try { body = (string)mail.Body; } catch { }
                        string baseDir = ResolveBaseDir(cfg, type, code);
                        _archiveBaseDir = baseDir;
                        _archiveCompany = company;
                        System.Collections.Generic.List<string> matches, existing, reasons; string newRel;
                        BuildArchiveSuggestions(baseDir, subject, sender, body, year, company, sap,
                                                out matches, out reasons, out newRel, out existing);
                        picker.SetFolders(existing.ToArray());
                        picker.SetSuggestions(matches.ToArray(), reasons.ToArray(), newRel);
                    }
                    catch { }
                });
                worker.IsBackground = true; worker.Start();

                try
                {
                    var r = picker.ShowDialog();
                    if (r == DialogResult.OK)
                    {
                        string rel = !string.IsNullOrEmpty(picker.CreateFolder) ? picker.CreateFolder : picker.Chosen;
                        if (!string.IsNullOrEmpty(rel))
                        {
                            string root = _archiveBaseDir ?? cfg.ClientBase;
                            string abs = Path.IsPathRooted(rel) ? rel : Path.Combine(root, rel);
                            if (!string.IsNullOrWhiteSpace(cfg.Subfolder)) abs = Path.Combine(abs, cfg.Subfolder);
                            SaveEmailPerMode(mail, abs, subject, cfg.SaveMode);
                            SaveArchiveChoice(_archiveCompany, rel);   // learn where this client's mail goes
                        }
                    }
                }
                finally { picker.Dispose(); }
            }
            catch (Exception ex) { Ui.Notify("Axon error: " + ex.Message, "Axon intelligence"); }
        }

        // ---- Email-archive: read config, extract entities, suggest folders, save per mode ----
        private class ArchiveCfg
        {
            public string ClientBase = "", SupplierBase = "", SaveMode = "both", Subfolder = "";
            public System.Collections.Generic.Dictionary<string, string> Codes =
                new System.Collections.Generic.Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            public bool Ready { get { return !string.IsNullOrWhiteSpace(ClientBase); } }
        }

        private ArchiveCfg ReadArchiveCfg()
        {
            var cfg = new ArchiveCfg();
            try
            {
                string p = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
                                        "AxonOutlook", "archive.json");
                if (!File.Exists(p)) return cfg;
                var js = new System.Web.Script.Serialization.JavaScriptSerializer();
                var d = js.DeserializeObject(File.ReadAllText(p)) as System.Collections.Generic.Dictionary<string, object>;
                if (d == null) return cfg;
                if (d.ContainsKey("client_base") && d["client_base"] != null) cfg.ClientBase = d["client_base"].ToString();
                if (d.ContainsKey("supplier_base") && d["supplier_base"] != null) cfg.SupplierBase = d["supplier_base"].ToString();
                if (d.ContainsKey("save_mode") && d["save_mode"] != null) cfg.SaveMode = d["save_mode"].ToString();
                if (d.ContainsKey("default_subfolder") && d["default_subfolder"] != null) cfg.Subfolder = d["default_subfolder"].ToString();
                var cc = d.ContainsKey("country_codes") ? d["country_codes"] as System.Collections.Generic.Dictionary<string, object> : null;
                if (cc != null) foreach (var kv in cc) if (kv.Value != null) cfg.Codes[kv.Key] = kv.Value.ToString();
            }
            catch { }
            return cfg;
        }

        private System.Collections.Generic.Dictionary<string, object> ExtractArchiveInfo(dynamic mail, ArchiveCfg cfg)
        {
            string subject = ""; try { subject = (string)mail.Subject; } catch { }
            string sender = ""; try { sender = (string)mail.SenderName; } catch { }
            string senderEmail = ""; try { senderEmail = (string)mail.SenderEmailAddress; } catch { }
            string body = ""; try { body = (string)mail.Body; } catch { }
            if (body.Length > 1500) body = body.Substring(0, 1500);
            var js = new System.Web.Script.Serialization.JavaScriptSerializer();
            string mapJson = js.Serialize(cfg.Codes);
            string prompt =
                "Read this email and reply with ONLY JSON: {\"code\":\"\",\"company\":\"\",\"type\":\"client|supplier\"," +
                "\"year\":\"\",\"sap\":\"\"}. Determine the sender's COUNTRY from the email domain, any phone numbers " +
                "(+32 Belgium, +49 Germany, +31 Netherlands, +33 France, etc.), and address, then set \"code\" using this " +
                "Country->code map: " + mapJson + " (empty if the country isn't in the map). \"company\" = the client/supplier " +
                "company name. \"type\" = 'client' if they buy from us (an order to us) or 'supplier' if they sell to us " +
                "(their quote/invoice). \"year\" = a 4-digit year from the email, else the current year. \"sap\" = the order/" +
                "SAP number in the subject if there is one, else empty.\n\nFrom: " + sender + " <" + senderEmail + ">\n" +
                "Subject: " + subject + "\n\n" + body;
            string text = ModelComplete(prompt, 0);
            var m = System.Text.RegularExpressions.Regex.Match(text ?? "", "\\{[\\s\\S]*\\}");
            if (!m.Success) return null;
            try { return js.DeserializeObject(m.Value) as System.Collections.Generic.Dictionary<string, object>; }
            catch { return null; }
        }

        private string ResolveBaseDir(ArchiveCfg cfg, string type, string code)
        {
            string bas = ((type == "supplier") && !string.IsNullOrWhiteSpace(cfg.SupplierBase)) ? cfg.SupplierBase : cfg.ClientBase;
            bas = bas.TrimEnd('\\', '/');
            if (!string.IsNullOrWhiteSpace(code))
            {
                var codeVals = new System.Collections.Generic.HashSet<string>(cfg.Codes.Values, StringComparer.OrdinalIgnoreCase);
                var parts = bas.Split('\\');
                for (int i = 0; i < parts.Length; i++)
                    if (codeVals.Contains(parts[i])) { parts[i] = code; break; }
                bas = string.Join("\\", parts);
            }
            return bas;
        }

        private void BuildArchiveSuggestions(string baseDir, string subject, string sender, string body,
            string year, string company, string sap,
            out System.Collections.Generic.List<string> matches, out System.Collections.Generic.List<string> reasons,
            out string newRel, out System.Collections.Generic.List<string> existing)
        {
            var mm = new System.Collections.Generic.List<string>();   // locals so the lambda can capture them
            var rr = new System.Collections.Generic.List<string>();
            matches = mm;
            reasons = rr;
            existing = new System.Collections.Generic.List<string>();
            var parts = new System.Collections.Generic.List<string>();
            if (!string.IsNullOrWhiteSpace(year)) parts.Add(year);
            if (!string.IsNullOrWhiteSpace(company)) parts.Add(company);
            if (!string.IsNullOrWhiteSpace(sap)) parts.Add(sap);
            newRel = string.Join("\\", parts);

            Action<string, string> add = (rel, why) =>
            {
                if (string.IsNullOrEmpty(rel) || mm.Count >= 5) return;
                foreach (var m in mm) if (string.Equals(m, rel, StringComparison.OrdinalIgnoreCase)) return;
                mm.Add(rel); rr.Add(why);
            };

            try
            {
                if (!Directory.Exists(baseDir)) return;
                CollectDirs(baseDir, baseDir, 0, 3, existing);
                var existingSet = new System.Collections.Generic.HashSet<string>(existing, StringComparer.OrdinalIgnoreCase);

                // #1 Memory: if this client's mail was filed somewhere before and it still exists, lead with it.
                string remembered = RememberedFolder(company);
                if (!string.IsNullOrEmpty(remembered) && existingSet.Contains(remembered))
                    add(remembered, "filed here before");

                // Preferred: let the AI pick where THIS email belongs (understands client + topic).
                string aiNew;
                System.Collections.Generic.List<string> aiReasons;
                var aiMatches = RankArchiveViaApi(subject, sender, body, company, existing, out aiReasons, out aiNew);
                if (aiMatches != null && aiMatches.Count > 0)
                {
                    for (int i = 0; i < aiMatches.Count; i++)
                        add(aiMatches[i], i < aiReasons.Count ? aiReasons[i] : "best match");
                    if (!string.IsNullOrWhiteSpace(aiNew)) newRel = aiNew;
                    if (matches.Count > 0) return;
                }

                // Fallback: keyword scoring. Require a real signal (SAP or client name); a bare year match
                // is too weak on its own, so it no longer floods the list with every "2025" folder.
                var scored = new System.Collections.Generic.List<System.Collections.Generic.KeyValuePair<int, string>>();
                foreach (var rel in existing)
                {
                    int s = 0;
                    if (!string.IsNullOrEmpty(sap) && rel.IndexOf(sap, StringComparison.OrdinalIgnoreCase) >= 0) s += 100;
                    if (!string.IsNullOrEmpty(company) && rel.IndexOf(company, StringComparison.OrdinalIgnoreCase) >= 0) s += 10;
                    if (!string.IsNullOrEmpty(year) && rel.IndexOf(year, StringComparison.Ordinal) >= 0) s += 1;
                    if (s >= 10) scored.Add(new System.Collections.Generic.KeyValuePair<int, string>(s, rel));
                }
                scored.Sort((a, b) => b.Key.CompareTo(a.Key));
                foreach (var kv in scored) add(kv.Value, kv.Key >= 100 ? "order number" : "client match");
            }
            catch { }
        }

        // Ask the model to choose the best archive folders for this email from the ones that exist.
        // Returns exact-matching relative paths (validated against `existing` so it can't hallucinate),
        // a 1-2 word reason per pick, and an optional new folder that follows the same structure.
        private System.Collections.Generic.List<string> RankArchiveViaApi(string subject, string sender,
            string body, string company, System.Collections.Generic.List<string> existing,
            out System.Collections.Generic.List<string> reasons, out string newFolder)
        {
            newFolder = "";
            reasons = new System.Collections.Generic.List<string>();
            try
            {
                var js = new System.Web.Script.Serialization.JavaScriptSerializer();
                // Keep the list a sane size: the client's own folders first, then a sample of the rest.
                var list = new System.Collections.Generic.List<string>();
                if (!string.IsNullOrEmpty(company))
                    foreach (var r in existing)
                        if (r.IndexOf(company, StringComparison.OrdinalIgnoreCase) >= 0) list.Add(r);
                foreach (var r in existing) { if (list.Count >= 350) break; if (!list.Contains(r)) list.Add(r); }

                string b = body ?? ""; if (b.Length > 1200) b = b.Substring(0, 1200);
                string prompt =
                    "You are filing an email into one of the user's ARCHIVE folders on disk (organised by client, " +
                    "topic and year).\n\nEmail\n  Subject: " + subject + "\n  From: " + sender +
                    "\n  Body (truncated):\n" + b + "\n\nClient/company (best guess): " + company + "\n\n" +
                    "The user's existing archive folders (relative paths):\n" + js.Serialize(list.ToArray()) + "\n\n" +
                    "Decide where THIS email should be saved, based on the client/company and what the email is ABOUT " +
                    "(e.g. the platform, campaign or project named in the subject). List the best-fitting folders first " +
                    "(up to 5), copied EXACTLY from the list, and ONLY genuinely good fits (do not pad). " +
                    "Prefer the DEEPEST, most specific folder for this client (its own subfolder, and the right year " +
                    "subfolder) over a generic parent folder. If nothing fits well, propose a short NEW folder path that " +
                    "follows the SAME structure as the existing folders.\n" +
                    "For each pick give a 1-2 word reason (e.g. 'client', 'same topic', 'client + year').\n" +
                    "Reply with ONLY JSON: {\"matches\":[{\"path\":\"exact folder path\",\"why\":\"1-2 words\"}], " +
                    "\"new_folder\":\"a relative path or empty\"}";
                string text = ModelComplete(prompt, 0);
                if (string.IsNullOrEmpty(text)) return null;
                var mt = System.Text.RegularExpressions.Regex.Match(text, "\\{[\\s\\S]*\\}");
                if (!mt.Success) return null;
                var d = js.DeserializeObject(mt.Value) as System.Collections.Generic.Dictionary<string, object>;
                if (d == null) return null;
                var outList = new System.Collections.Generic.List<string>();
                var existingSet = new System.Collections.Generic.HashSet<string>(existing, StringComparer.OrdinalIgnoreCase);
                if (d.ContainsKey("matches") && d["matches"] is object[])
                    foreach (var o in (object[])d["matches"])
                    {
                        string val = "", why = "best match";
                        var row = o as System.Collections.Generic.Dictionary<string, object>;
                        if (row != null)
                        {
                            if (row.ContainsKey("path") && row["path"] != null) val = row["path"].ToString();
                            if (row.ContainsKey("why") && row["why"] != null) why = row["why"].ToString();
                        }
                        else { val = (o == null ? "" : o.ToString()); }   // tolerate a plain string
                        val = val.Replace("/", "\\").Trim();
                        if (existingSet.Contains(val) && !outList.Contains(val)) { outList.Add(val); reasons.Add(why.Trim()); }
                        if (outList.Count >= 5) break;
                    }
                if (d.ContainsKey("new_folder") && d["new_folder"] != null)
                    newFolder = d["new_folder"].ToString().Replace("/", "\\").Trim();
                return outList;
            }
            catch { newFolder = ""; reasons = new System.Collections.Generic.List<string>(); return null; }
        }

        private static void CollectDirs(string root, string dir, int depth, int maxDepth,
            System.Collections.Generic.List<string> outList)
        {
            if (depth >= maxDepth || outList.Count > 600) return;
            try
            {
                foreach (var d in Directory.GetDirectories(dir))
                {
                    outList.Add(d.Substring(root.Length).TrimStart('\\', '/'));
                    CollectDirs(root, d, depth + 1, maxDepth, outList);
                    if (outList.Count > 600) return;
                }
            }
            catch { }
        }

        private void SaveEmailPerMode(dynamic mail, string folder, string subject, string mode)
        {
            try
            {
                Directory.CreateDirectory(folder);
                mode = (mode ?? "both").ToLowerInvariant();
                bool saveMsg = mode == "both" || mode == "email";
                bool saveAtt = mode == "both" || mode == "attachments";
                if (saveMsg)
                {
                    string name = string.IsNullOrEmpty(subject) ? "email" : subject;
                    foreach (char c in Path.GetInvalidFileNameChars()) name = name.Replace(c, ' ');
                    name = name.Trim(); if (name.Length == 0) name = "email"; if (name.Length > 120) name = name.Substring(0, 120);
                    string path = Path.Combine(folder, name + ".msg");
                    int i = 1;
                    while (File.Exists(path)) { path = Path.Combine(folder, name + " (" + i + ").msg"); i++; }
                    mail.SaveAs(path, 9);   // olMSGUnicode
                }
                if (saveAtt)
                {
                    try
                    {
                        foreach (dynamic att in mail.Attachments)
                        {
                            try
                            {
                                string an = (string)att.FileName;
                                foreach (char c in Path.GetInvalidFileNameChars()) an = an.Replace(c, ' ');
                                string ap = Path.Combine(folder, an);
                                string bn = Path.GetFileNameWithoutExtension(ap), ex = Path.GetExtension(ap);
                                int j = 1;
                                while (File.Exists(ap)) { ap = Path.Combine(folder, bn + " (" + j + ")" + ex); j++; }
                                att.SaveAsFile(ap);
                            }
                            catch { }
                        }
                    }
                    catch { }
                }
                Ui.Notify("Saved to:\n" + folder, "Axon intelligence");
            }
            catch (Exception ex) { Ui.Notify("Couldn't save: " + ex.Message, "Axon intelligence"); }
        }

        // Save the email as a .msg into the chosen disk folder (unique filename from the subject).
        private void SaveEmail(dynamic mail, string folder, string subject)
        {
            try
            {
                string name = string.IsNullOrEmpty(subject) ? "email" : subject;
                foreach (char c in Path.GetInvalidFileNameChars()) name = name.Replace(c, ' ');
                name = name.Trim();
                if (name.Length == 0) name = "email";
                if (name.Length > 120) name = name.Substring(0, 120);
                string path = Path.Combine(folder, name + ".msg");
                int i = 1;
                while (File.Exists(path)) { path = Path.Combine(folder, name + " (" + i + ").msg"); i++; }
                mail.SaveAs(path, 9);   // 9 = olMSGUnicode
                Ui.Notify("Saved to:\n" + path, "Axon intelligence");
            }
            catch (Exception ex) { Ui.Notify("Couldn't save: " + ex.Message, "Axon intelligence"); }
        }

        // The user's configured save-folders (disk paths) live in %APPDATA%\AxonIntelligence.
        private static string DownloadConfigPath()
        {
            string dir = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "AxonIntelligence");
            try { Directory.CreateDirectory(dir); } catch { }
            return Path.Combine(dir, "download_folders.json");
        }

        private System.Collections.Generic.List<string> LoadDownloadFolders()
        {
            var list = new System.Collections.Generic.List<string>();
            try
            {
                string p = DownloadConfigPath();
                if (File.Exists(p))
                {
                    var js = new System.Web.Script.Serialization.JavaScriptSerializer();
                    var arr = js.DeserializeObject(File.ReadAllText(p)) as object[];
                    if (arr != null) foreach (var x in arr) if (x != null) list.Add(x.ToString());
                }
            }
            catch { }
            return list;
        }

        private void SaveDownloadFolders(System.Collections.Generic.List<string> folders)
        {
            try
            {
                var js = new System.Web.Script.Serialization.JavaScriptSerializer();
                File.WriteAllText(DownloadConfigPath(), js.Serialize(folders), new System.Text.UTF8Encoding(false));
            }
            catch { }
        }

        // The currently open or selected mail item (Class 43 = olMail), or null.
        private object GetSelectedMail()
        {
            dynamic app = _app;
            if (app == null) return null;
            try
            {
                dynamic insp = app.ActiveInspector();
                if (insp != null) { dynamic item = insp.CurrentItem; if (item != null && (int)item.Class == 43) return item; }
            }
            catch { }
            try
            {
                dynamic exp = app.ActiveExplorer();
                if (exp != null) { dynamic sel = exp.Selection; if (sel != null && (int)sel.Count >= 1) { dynamic item = sel.Item(1); if ((int)item.Class == 43) return item; } }
            }
            catch { }
            return null;
        }

        public void OnSummarize(object control)
        {
            try
            {
                object m = GetSelectedMail();
                if (m == null) { Ui.Notify("Select an email first.", "Axon intelligence"); return; }
                dynamic mail = m;
                string subj = ""; try { subj = (string)mail.Subject; } catch { }
                string body = ""; try { body = (string)mail.Body; } catch { }
                if (body.Length > 6000) body = body.Substring(0, 6000);
                string bodyCopy = body;
                bool wantsReply;
                using (var dlg = new SummaryDialog(subj,
                    lang => ModelComplete(BuildSummaryPrompt(subj, bodyCopy, lang), 0.3)))
                {
                    dlg.ShowDialog();
                    wantsReply = dlg.WantsReply;
                }
                if (wantsReply) OnReply(control);   // jump straight into the reply flow on the same email
            }
            catch (Exception ex) { Ui.Notify("Axon error: " + ex.Message, "Axon intelligence"); }
        }

        private string BuildSummaryPrompt(string subject, string body, string lang)
        {
            string langLine = string.IsNullOrEmpty(lang)
                ? "Write the summary in the same language as the email."
                : "Write the summary in " + lang + ".";
            return "Summarize this email for a busy reader in PLAIN TEXT only — no markdown, no asterisks, " +
                "no '#'. Use exactly this layout, with blank lines between sections:\n\nGist: <one sentence>\n\n" +
                "Key points:\n- <point>\n- <point>\n\nAction: <what the reader should do, or 'None'>\n\n" +
                langLine + " Keep it tight.\n\nSubject: " + subject + "\n\n" + body;
        }

        public void OnReply(object control)
        {
            try
            {
                object m = GetSelectedMail();
                if (m == null) { Ui.Notify("Select an email first.", "Axon intelligence"); return; }
                dynamic mail = m;
                string subj = ""; try { subj = (string)mail.Subject; } catch { }
                string sender = ""; try { sender = (string)mail.SenderName; } catch { }
                string body = ""; try { body = (string)mail.Body; } catch { }
                if (body.Length > 6000) body = body.Substring(0, 6000);
                string me = ""; try { me = (string)((dynamic)_app).Session.CurrentUser.Name; } catch { }
                string bodyCopy = body;
                // Ask the user HOW they want to reply, then Axon drafts to that instruction.
                string draft;
                using (var prompt = new ReplyPrompt(subj,
                    (instr, lang) => ModelComplete(BuildReplyPrompt(subj, sender, bodyCopy, instr, me, lang), 0.4)))
                {
                    if (prompt.ShowDialog() != DialogResult.OK) return;
                    draft = prompt.Draft;
                }
                if (string.IsNullOrEmpty(draft)) { Ui.Notify("Couldn't draft a reply (model unavailable).", "Axon intelligence"); return; }
                dynamic reply = mail.Reply();
                // Draft at the top, a horizontal line, then Outlook's quoted original (keeps formatting).
                try
                {
                    string html = (string)reply.HTMLBody;
                    // No <hr> here — Outlook's reply already has its own divider above the quoted original.
                    string draftHtml = "<div style='font-family:Calibri,sans-serif;font-size:11pt;'>" +
                        ToHtml(draft) + "<br></div>";
                    int bi = html.IndexOf("<body", StringComparison.OrdinalIgnoreCase);
                    int gt = bi >= 0 ? html.IndexOf('>', bi) : -1;
                    reply.HTMLBody = gt >= 0 ? html.Substring(0, gt + 1) + draftHtml + html.Substring(gt + 1)
                                             : draftHtml + html;
                }
                catch { try { reply.Body = draft; } catch { } }
                reply.Display();   // open the draft for the user to review + send (never auto-sent)
            }
            catch (Exception ex) { Ui.Notify("Axon error: " + ex.Message, "Axon intelligence"); }
        }

        // Plain text -> minimal HTML (escape + line breaks) for inserting into the reply.
        private static string ToHtml(string text)
        {
            string e = (text ?? "").Replace("&", "&amp;").Replace("<", "&lt;").Replace(">", "&gt;");
            return e.Replace("\r\n", "\n").Replace("\n", "<br>");
        }

        private string BuildReplyPrompt(string subject, string sender, string body, string instruction, string me, string lang)
        {
            string how = string.IsNullOrWhiteSpace(instruction)
                ? "Write a concise, professional, courteous reply that addresses the points raised."
                : "Write the reply following these instructions from the user: " + instruction;
            string langLine = string.IsNullOrEmpty(lang)
                ? "Write the reply in the SAME language as the email."
                : "Write the ENTIRE reply (greeting, message, and sign-off) in " + lang + ", regardless of the email's language.";
            string tone = MyToneGuide();
            string toneLine = string.IsNullOrEmpty(tone) ? "" : " Match the user's personal writing style:\n" + tone + "\n";
            return how + " Begin with an appropriate greeting addressed to the sender by first name, and end " +
                   "with a courteous sign-off" + (string.IsNullOrWhiteSpace(me) ? "" : " from " + me) + ". " +
                   langLine + toneLine + " Use a natural tone. Output ONLY the reply text itself (greeting, " +
                   "message, sign-off) — no subject line and no quoted original.\n\n" +
                   "Sender: " + sender + "\nSubject: " + subject + "\n\n" + body;
        }

        // The user's learned writing style (saved by the dot's "learn my tone" to %APPDATA%\AxonOutlook\tone.txt).
        private static string MyToneGuide()
        {
            try
            {
                string p = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
                                        "AxonOutlook", "tone.txt");
                if (File.Exists(p)) { string s = File.ReadAllText(p).Trim(); if (s.Length > 0) return s; }
            }
            catch { }
            return "";
        }

        public void OnSchedule(object control)
        {
            try
            {
                object m = GetSelectedMail();
                if (m == null) { Ui.Notify("Select an email first.", "Axon intelligence"); return; }
                dynamic mail = m;
                string subj = ""; try { subj = (string)mail.Subject; } catch { }
                string sender = ""; try { sender = (string)mail.SenderName; } catch { }
                string body = ""; try { body = (string)mail.Body; } catch { }
                if (body.Length > 6000) body = body.Substring(0, 6000);
                // Ask the user WHEN (any timeframe — "next Tue 2pm", "week of 15 Sept", "in 3 months",
                // or blank to suggest soon) and how long. Axon then drafts the invite to match.
                string when; int dur;
                using (var sp = new SchedulePrompt(subj))
                {
                    if (sp.ShowDialog() != DialogResult.OK) return;
                    when = sp.When; dur = sp.DurationMinutes;
                }
                string today = DateTime.Now.ToString("yyyy-MM-dd (dddd)");
                string busy = MyBusyBlocks((dynamic)_app, 14);   // next 2 weeks, for near-term conflict checks
                string text = ModelComplete(BuildSchedulePrompt(subj, sender, body, today, busy, when), 0.2);
                if (string.IsNullOrEmpty(text)) { Ui.Notify("Couldn't propose a meeting (model unavailable).", "Axon intelligence"); return; }
                var mt = System.Text.RegularExpressions.Regex.Match(text, "\\{[\\s\\S]*\\}");
                if (!mt.Success) { Ui.Notify("Couldn't read the proposed meeting details.", "Axon intelligence"); return; }
                var js = new System.Web.Script.Serialization.JavaScriptSerializer();
                var d = (System.Collections.Generic.Dictionary<string, object>)js.DeserializeObject(mt.Value);

                string mSubject = Field(d, "subject");
                if (string.IsNullOrWhiteSpace(mSubject)) mSubject = "Meeting: " + subj;
                DateTime start;
                if (!DateTime.TryParse(Field(d, "start"), System.Globalization.CultureInfo.InvariantCulture,
                        System.Globalization.DateTimeStyles.None, out start))
                    start = DateTime.Now.Date.AddDays(1).AddHours(10);
                string loc = Field(d, "location");
                string agenda = Field(d, "agenda");

                dynamic app = _app;
                dynamic appt = app.CreateItem(1);   // olAppointmentItem
                appt.MeetingStatus = 1;             // olMeeting -> a sendable invite
                appt.Subject = mSubject;
                appt.Start = start;
                appt.Duration = dur;
                if (!string.IsNullOrEmpty(loc)) appt.Location = loc;
                if (!string.IsNullOrEmpty(agenda)) appt.Body = agenda;
                // Invite everyone on the email: sender + To as Required, Cc as Optional (minus yourself).
                string meSmtp = ""; try { meSmtp = SmtpOf(app.Session.CurrentUser.AddressEntry); } catch { }
                var seen = new System.Collections.Generic.HashSet<string>(StringComparer.OrdinalIgnoreCase);
                if (!string.IsNullOrEmpty(meSmtp)) seen.Add(meSmtp);
                Action<string, int> addAttendee = (addr, type) =>
                {
                    if (string.IsNullOrEmpty(addr) || !seen.Add(addr)) return;
                    try { dynamic rr = appt.Recipients.Add(addr); rr.Type = type; } catch { }
                };
                try { addAttendee(SmtpOf(mail.Sender), 1); } catch { }            // sender -> Required
                try
                {
                    foreach (dynamic r in mail.Recipients)
                    {
                        string a = ""; try { a = SmtpOf(r.AddressEntry); } catch { }
                        if (string.IsNullOrEmpty(a)) { try { a = (string)r.Address; } catch { } }
                        int t = 1; try { t = (int)r.Type; } catch { }            // mail: 1=To, 2=Cc, 3=Bcc
                        addAttendee(a, t == 2 ? 2 : 1);                           // Cc -> Optional, else Required
                    }
                }
                catch { }
                try { appt.Recipients.ResolveAll(); } catch { }
                // Single clean window: the pre-filled invite (already at a free time). Outlook's own
                // Scheduling Assistant inside it shows everyone's free/busy if the user wants to adjust.
                appt.Display();   // review + send (never auto-sent)
            }
            catch (Exception ex) { Ui.Notify("Axon error: " + ex.Message, "Axon intelligence"); }
        }

        private static string Field(System.Collections.Generic.Dictionary<string, object> d, string k)
        {
            return (d != null && d.ContainsKey(k) && d[k] != null) ? d[k].ToString() : "";
        }

        // Resolve a sender/recipient AddressEntry to its SMTP address (handles Exchange senders).
        private static string SmtpOf(dynamic ae)
        {
            try
            {
                if (ae == null) return "";
                try { dynamic eu = ae.GetExchangeUser(); if (eu != null) { string s = (string)eu.PrimarySmtpAddress; if (!string.IsNullOrEmpty(s)) return s; } } catch { }
                try { dynamic pa = ae.PropertyAccessor; string s = (string)pa.GetProperty("http://schemas.microsoft.com/mapi/proptag/0x39FE001F"); if (!string.IsNullOrEmpty(s)) return s; } catch { }
                try { string s = (string)ae.Address; if (!string.IsNullOrEmpty(s)) return s; } catch { }
            }
            catch { }
            return "";
        }

        private string BuildSchedulePrompt(string subject, string sender, string body, string today, string busy, string when)
        {
            string target = string.IsNullOrWhiteSpace(when)
                ? "Pick the soonest suitable time during business hours (Mon–Fri, 09:00–17:00)."
                : "The user wants the meeting around: \"" + when + "\". Resolve that relative to today into a " +
                  "concrete date and time during business hours (Mon–Fri, 09:00–17:00).";
            string avail = string.IsNullOrEmpty(busy)
                ? ""
                : " If the chosen time falls within the next 2 weeks, do NOT overlap these busy blocks:\n" + busy;
            return "Read this email and propose a meeting with the sender. Today is " + today + ". " + target +
                avail + "\nIf the email itself proposes a specific time and it fits, prefer that.\n\n" +
                "Reply with ONLY JSON:\n{\"subject\": \"a clear meeting title\", " +
                "\"start\": \"YYYY-MM-DDTHH:MM\" (24h), \"location\": \"a place, or 'Microsoft Teams', or empty\", " +
                "\"agenda\": \"a short 1-3 line agenda\"}.\n\n" +
                "From: " + sender + "\nSubject: " + subject + "\n\n" + body;
        }

        // The user's busy blocks over the next `days` days (recurrences expanded), for conflict-free
        // scheduling. Best-effort: returns "" if the calendar can't be read.
        private static string MyBusyBlocks(dynamic app, int days)
        {
            try
            {
                dynamic items = app.Session.GetDefaultFolder(9).Items;   // 9 = olFolderCalendar
                items.IncludeRecurrences = true;
                items.Sort("[Start]");
                DateTime from = DateTime.Now;
                DateTime to = from.AddDays(days);
                var ci = System.Globalization.CultureInfo.CurrentCulture;
                string filter = "[Start] < '" + to.ToString("g", ci) + "' AND [End] > '" + from.ToString("g", ci) + "'";
                var sb = new System.Text.StringBuilder();
                int n = 0;
                foreach (dynamic a in items.Restrict(filter))
                {
                    try
                    {
                        DateTime s = (DateTime)a.Start;
                        DateTime e = (DateTime)a.End;
                        bool allday = false; try { allday = (bool)a.AllDayEvent; } catch { }
                        string subj = ""; try { subj = (string)a.Subject; } catch { }
                        sb.AppendLine("- " + s.ToString("ddd yyyy-MM-dd HH:mm") + "-" + e.ToString("HH:mm") +
                            (allday ? " (all day)" : "") + (string.IsNullOrEmpty(subj) ? "" : " : " + subj));
                    }
                    catch { }
                    if (++n >= 60) break;
                }
                return sb.ToString().Trim();
            }
            catch { return ""; }
        }

        public void OnAttachEmail(object control)
        {
            try
            {
                dynamic app = _app;
                var items = new System.Collections.Generic.List<dynamic>();
                // Prefer the current selection (may be several); fall back to a single open email.
                try
                {
                    dynamic exp = app.ActiveExplorer();
                    if (exp != null && exp.Selection != null)
                    {
                        foreach (dynamic it in exp.Selection)
                        {
                            try { if ((int)it.Class == 43) items.Add(it); } catch { }
                        }
                    }
                }
                catch { }
                if (items.Count == 0)
                {
                    object m = GetSelectedMail();
                    if (m != null) items.Add((dynamic)m);
                }
                if (items.Count == 0) { Ui.Notify("Select an email first.", "Axon intelligence"); return; }

                dynamic nm = app.CreateItem(0);   // olMailItem
                string firstSubj = "";
                foreach (dynamic it in items)
                {
                    try { nm.Attachments.Add(it); } catch { }   // attach the message as a .msg
                    if (firstSubj == "") { try { firstSubj = (string)it.Subject; } catch { } }
                }
                nm.Subject = items.Count > 1 ? ("FW: " + items.Count + " emails") : ("FW: " + firstSubj);
                nm.Display();   // open the new email for the user to add a recipient/note and send
            }
            catch (Exception ex) { Ui.Notify("Axon error: " + ex.Message, "Axon intelligence"); }
        }

        public void OnWriteEmail(object control)
        {
            try
            {
                dynamic app = _app;
                dynamic insp = app.ActiveInspector();
                if (insp == null) { Ui.Notify("Open a new email first, then click Write with Axon.", "Axon intelligence"); return; }
                dynamic item = insp.CurrentItem;
                int cls = 0; try { cls = (int)item.Class; } catch { }
                if (item == null || cls != 43) { Ui.Notify("Write with Axon works on an email you're composing.", "Axon intelligence"); return; }
                string me = ""; try { me = (string)app.Session.CurrentUser.Name; } catch { }

                string draft = null;
                using (var prompt = new WritePrompt(
                    (instr, lang) => ModelComplete(BuildWritePrompt(instr, me, lang), 0.5)))
                {
                    if (prompt.ShowDialog() != DialogResult.OK) return;
                    draft = prompt.Draft;
                }
                if (string.IsNullOrEmpty(draft)) { Ui.Notify("Couldn't write the email (model unavailable).", "Axon intelligence"); return; }

                // Consume the leading "To:" / "Subject:" header lines; the rest is the body.
                string to = null, subject = null;
                var lines = draft.Replace("\r\n", "\n").Split('\n');
                int bstart = 0;
                for (; bstart < lines.Length; bstart++)
                {
                    string ln = lines[bstart].Trim();
                    if (ln.Length == 0) { if (to != null || subject != null) { bstart++; break; } continue; }
                    if (to == null && ln.StartsWith("To:", StringComparison.OrdinalIgnoreCase)) { to = ln.Substring(3).Trim(); continue; }
                    if (subject == null && ln.StartsWith("Subject:", StringComparison.OrdinalIgnoreCase)) { subject = ln.Substring(8).Trim(); continue; }
                    break;   // first non-header line -> body starts here
                }
                var bodySb = new System.Text.StringBuilder();
                for (int j = bstart; j < lines.Length; j++) bodySb.AppendLine(lines[j]);
                string body = bodySb.ToString().Trim();
                if (string.IsNullOrEmpty(body)) body = draft;   // fallback if no headers were found

                if (!string.IsNullOrEmpty(to))
                {
                    try
                    {
                        string curTo = (string)item.To;
                        if (string.IsNullOrWhiteSpace(curTo))
                        {
                            item.To = to;
                            try { item.Recipients.ResolveAll(); } catch { }   // resolve the name to an address
                        }
                    }
                    catch { }
                }
                if (!string.IsNullOrEmpty(subject))
                {
                    try { string cur = (string)item.Subject; if (string.IsNullOrWhiteSpace(cur)) item.Subject = subject; } catch { }
                }
                // Write the body into the open compose window (HTML) for the user to review + send.
                try
                {
                    string html = (string)item.HTMLBody;
                    string draftHtml = "<div style='font-family:Calibri,sans-serif;font-size:11pt;'>" + ToHtml(body) + "</div>";
                    int bi = html.IndexOf("<body", StringComparison.OrdinalIgnoreCase);
                    int gt = bi >= 0 ? html.IndexOf('>', bi) : -1;
                    item.HTMLBody = gt >= 0 ? html.Substring(0, gt + 1) + draftHtml + html.Substring(gt + 1) : draftHtml + html;
                }
                catch { try { item.Body = body; } catch { } }
            }
            catch (Exception ex) { Ui.Notify("Axon error: " + ex.Message, "Axon intelligence"); }
        }

        private string BuildWritePrompt(string instruction, string me, string lang)
        {
            string langLine = string.IsNullOrEmpty(lang)
                ? "Write the email in the same language the description is written in."
                : "Write the ENTIRE email in " + lang + ".";
            string tone = MyToneGuide();
            string toneLine = string.IsNullOrEmpty(tone) ? "" : " Match the user's personal writing style:\n" + tone + "\n";
            return "Write a complete, professional email based on this description from the user:\n" + instruction +
                   "\n\nBegin with an appropriate greeting and end with a courteous sign-off" +
                   (string.IsNullOrWhiteSpace(me) ? "" : " from " + me) + ". " + langLine + toneLine +
                   " At the very top, output these header lines (each on its own line): " +
                   "'To: <the recipient's name or email address if the description names one, otherwise leave " +
                   "it blank>' and 'Subject: <a short, clear subject>'. Then a blank line, then the email body " +
                   "(greeting addressed to the recipient by first name, message, sign-off). Output only that — no notes.";
        }

        public void OnSendLater(object control)
        {
            try
            {
                dynamic app = _app;
                dynamic insp = app.ActiveInspector();
                if (insp == null) { Ui.Notify("Open or start an email first, then use Send Later.", "Axon intelligence"); return; }
                dynamic item = insp.CurrentItem;
                int cls = 0; try { cls = (int)item.Class; } catch { }
                if (item == null || cls != 43) { Ui.Notify("Send Later works on an email you're composing.", "Axon intelligence"); return; }
                DateTime when; bool remind;
                using (var p = new WhenPrompt("Send Later", "When should Axon send this email?", true))
                {
                    if (p.ShowDialog() != DialogResult.OK) return;
                    when = p.When; remind = p.Remind;
                }
                if (when <= DateTime.Now.AddMinutes(1)) { Ui.Notify("Pick a time in the future.", "Axon intelligence"); return; }
                string subj = ""; try { subj = (string)item.Subject; } catch { }
                string to = ""; try { to = (string)item.To; } catch { }
                item.DeferredDeliveryTime = when;
                item.Send();   // moves to the Outbox and is sent automatically at that time
                WriteScheduled(subj, to, when, remind ? 10 : 0);   // dot notifies when it sends (+ heads-up)
                Ui.Notify("Scheduled — Axon will send this at " + when.ToString("ddd dd MMM, HH:mm") +
                          ". It waits in your Outbox until then (keep Outlook connected).", "Axon intelligence");
            }
            catch (Exception ex) { Ui.Notify("Axon error: " + ex.Message, "Axon intelligence"); }
        }

        public void OnFollowUp(object control)
        {
            try
            {
                object m = GetSelectedMail();
                if (m == null) { Ui.Notify("Select an email first.", "Axon intelligence"); return; }
                dynamic mail = m;
                DateTime when;
                using (var p = new WhenPrompt("Follow up", "When should Axon remind you to follow up?"))
                {
                    if (p.ShowDialog() != DialogResult.OK) return;
                    when = p.When;
                }
                try { mail.FlagRequest = "Follow up"; } catch { }
                try { mail.TaskStartDate = DateTime.Now; } catch { }
                try { mail.TaskDueDate = when; } catch { }
                try { mail.ReminderTime = when; mail.ReminderSet = true; } catch { }
                try { mail.FlagStatus = 2; } catch { }   // olFlagMarked (visual marker in Outlook)
                try { mail.Save(); } catch { }
                WriteFollowUp(mail, when);   // the dot pops a reliable reminder at 'when' (and skips it if replied)
                Ui.Notify("Follow-up set for " + when.ToString("ddd dd MMM, HH:mm") +
                          ". Axon will remind you then — unless they've already replied.", "Axon intelligence");
            }
            catch (Exception ex) { Ui.Notify("Axon error: " + ex.Message, "Axon intelligence"); }
        }

        // Record a follow-up so the always-running dot can pop a reliable reminder at the due time
        // (and skip it if a reply arrived). Shared file: %APPDATA%\AxonOutlook\followups.json.
        private void WriteFollowUp(dynamic mail, DateTime due)
        {
            try
            {
                string id = ""; try { id = (string)mail.EntryID; } catch { }
                string subject = ""; try { subject = (string)mail.Subject; } catch { }
                string who = ""; try { who = (string)mail.SenderName; } catch { }
                if (string.IsNullOrEmpty(who)) { try { who = (string)mail.To; } catch { } }
                string topic = ""; try { topic = (string)mail.ConversationTopic; } catch { }

                string dir = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "AxonOutlook");
                Directory.CreateDirectory(dir);
                string path = Path.Combine(dir, "followups.json");
                var js = new System.Web.Script.Serialization.JavaScriptSerializer();
                var items = new System.Collections.Generic.List<object>();
                try { if (File.Exists(path)) { var arr = js.DeserializeObject(File.ReadAllText(path)) as object[]; if (arr != null) items.AddRange(arr); } }
                catch { }
                items.Add(new System.Collections.Generic.Dictionary<string, object> {
                    { "id", id }, { "subject", subject }, { "who", who }, { "topic", topic },
                    { "due", due.ToString("yyyy-MM-ddTHH:mm:ss") },
                    { "created", DateTime.Now.ToString("yyyy-MM-ddTHH:mm:ss") },
                    { "notified", false },
                });
                File.WriteAllText(path, js.Serialize(items));
            }
            catch { }
        }

        // Record a scheduled ('Send Later') send so the always-running dot can notify the user when
        // its time arrives. Shared file: %APPDATA%\AxonOutlook\scheduled.json (append-only).
        private void WriteScheduled(string subject, string to, DateTime when, int remindBefore)
        {
            try
            {
                string dir = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "AxonOutlook");
                Directory.CreateDirectory(dir);
                string path = Path.Combine(dir, "scheduled.json");
                var js = new System.Web.Script.Serialization.JavaScriptSerializer();
                var items = new System.Collections.Generic.List<object>();
                try { if (File.Exists(path)) { var arr = js.DeserializeObject(File.ReadAllText(path)) as object[]; if (arr != null) items.AddRange(arr); } }
                catch { }
                var entry = new System.Collections.Generic.Dictionary<string, object> {
                    { "subject", subject ?? "" }, { "to", to ?? "" },
                    { "when", when.ToString("yyyy-MM-ddTHH:mm:ss") },
                    { "created", DateTime.Now.ToString("yyyy-MM-ddTHH:mm:ss") },
                };
                if (remindBefore > 0)
                    entry["remind_at"] = when.AddMinutes(-remindBefore).ToString("yyyy-MM-ddTHH:mm:ss");
                items.Add(entry);
                File.WriteAllText(path, js.Serialize(items));
            }
            catch { }
        }

        // --- COM (de)registration: add/remove the Outlook add-in registry entry ---
        private const string AddinKey = @"Software\Microsoft\Office\Outlook\AddIns\Axon.OutlookAddin";

        [ComRegisterFunction]
        public static void RegisterFunction(Type t)
        {
            using (var k = Registry.CurrentUser.CreateSubKey(AddinKey))
            {
                k.SetValue("FriendlyName", "Axon intelligence");
                k.SetValue("Description", "File and download emails with Axon intelligence");
                k.SetValue("LoadBehavior", 3, RegistryValueKind.DWord);
            }
        }

        [ComUnregisterFunction]
        public static void UnregisterFunction(Type t)
        {
            try { Registry.CurrentUser.DeleteSubKey(AddinKey, false); } catch { }
        }
    }

    // Shared look-and-feel for the Axon dialogs (black & white, flat, Segoe UI).
    internal static class Ui
    {
        public static readonly System.Drawing.Color Ink = System.Drawing.Color.FromArgb(51, 51, 58);     // softened text (dark grey, not pure black)
        public static readonly System.Drawing.Color AccentBg = System.Drawing.Color.FromArgb(74, 74, 84);   // button fill — dark grey
        public static readonly System.Drawing.Color AccentHover = System.Drawing.Color.FromArgb(92, 92, 104);
        public static readonly System.Drawing.Color Muted = System.Drawing.Color.FromArgb(130, 130, 140);
        public static readonly System.Drawing.Color SelBg = System.Drawing.Color.FromArgb(236, 238, 248);
        public static readonly System.Drawing.Color Line = System.Drawing.Color.FromArgb(220, 220, 226);

        public static Label Title(string t, int x, int y, int w)
        { return new Label { Text = t, Left = x, Top = y, Width = w, Height = 26, ForeColor = Ink, Font = new System.Drawing.Font("Segoe UI", 13F, System.Drawing.FontStyle.Bold) }; }
        public static Label Sub(string t, int x, int y, int w)
        { return new Label { Text = t, Left = x, Top = y, Width = w, Height = 20, ForeColor = Muted, AutoEllipsis = true }; }
        public static Label Caption(string t, int x, int y, int w)
        { return new Label { Text = t, Left = x, Top = y, Width = w, Height = 15, ForeColor = Muted, Font = new System.Drawing.Font("Segoe UI", 8F, System.Drawing.FontStyle.Bold) }; }
        public static Label Hint(string t, int x, int y, int w, int h)
        { return new Label { Text = t, Left = x, Top = y, Width = w, Height = h, ForeColor = Muted }; }

        public static Button Accent(string text, int x, int y, int w, int h)
        {
            var b = new Button { Text = text, Left = x, Top = y, Width = w, Height = h, FlatStyle = FlatStyle.Flat, BackColor = AccentBg, ForeColor = System.Drawing.Color.White, Font = new System.Drawing.Font("Segoe UI", 9.5F, System.Drawing.FontStyle.Bold), Cursor = Cursors.Hand };
            b.FlatAppearance.BorderSize = 0;
            b.FlatAppearance.MouseOverBackColor = AccentHover;
            return b;
        }
        public static Button Subtle(string text, int x, int y, int w, int h)
        {
            var b = new Button { Text = text, Left = x, Top = y, Width = w, Height = h, FlatStyle = FlatStyle.Flat, BackColor = System.Drawing.Color.White, ForeColor = System.Drawing.Color.FromArgb(40, 40, 40), Font = new System.Drawing.Font("Segoe UI", 9.5F), Cursor = Cursors.Hand };
            b.FlatAppearance.BorderColor = Line;
            b.FlatAppearance.BorderSize = 1;
            return b;
        }
        // Full-width left-aligned accent button used for suggestion rows.
        public static Button RowBtn(string text, int x, int y, int w, int h)
        {
            var b = Accent(text, x, y, w, h);
            b.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            b.Padding = new Padding(14, 0, 0, 0);
            return b;
        }
        public static string Leaf(string p)
        {
            if (string.IsNullOrEmpty(p)) return p;
            int i = p.LastIndexOf(" / ", StringComparison.Ordinal);
            if (i >= 0) return p.Substring(i + 3);
            string t = p.TrimEnd('\\', '/');
            int j = Math.Max(t.LastIndexOf('\\'), t.LastIndexOf('/'));
            return j >= 0 ? t.Substring(j + 1) : t;
        }
        // Owner-draw a folder list row as: leaf name (bold) + full path (gray).
        public static void DrawFolderItem(ListBox list, DrawItemEventArgs e)
        {
            if (e.Index < 0) return;
            string path = list.Items[e.Index].ToString();
            bool sel = (e.State & DrawItemState.Selected) != 0;
            using (var bg = new System.Drawing.SolidBrush(sel ? SelBg : System.Drawing.Color.White))
                e.Graphics.FillRectangle(bg, e.Bounds);
            using (var nf = new System.Drawing.Font("Segoe UI", 10F, System.Drawing.FontStyle.Bold))
            using (var pf = new System.Drawing.Font("Segoe UI", 8F))
            using (var ink = new System.Drawing.SolidBrush(Ink))
            using (var mut = new System.Drawing.SolidBrush(Muted))
            {
                e.Graphics.DrawString(Leaf(path), nf, ink, e.Bounds.Left + 12, e.Bounds.Top + 6);
                e.Graphics.DrawString(path, pf, mut, e.Bounds.Left + 12, e.Bounds.Top + 26);
            }
        }

        // Branded replacement for MessageBox — matches the Axon dialog style.
        public static void Notify(string message, string title = "Axon intelligence")
        {
            using (var f = new Form())
            {
                f.Text = title;
                f.FormBorderStyle = FormBorderStyle.FixedDialog;
                f.StartPosition = FormStartPosition.CenterScreen;
                f.MaximizeBox = false; f.MinimizeBox = false; f.ShowInTaskbar = false;
                f.BackColor = System.Drawing.Color.White;
                f.Font = new System.Drawing.Font("Segoe UI", 9.5F);
                f.ClientSize = new System.Drawing.Size(400, 170);
                int W = f.ClientSize.Width, H = f.ClientSize.Height;
                f.Controls.Add(Title("Axon intelligence", 22, 18, W - 44));
                f.Controls.Add(new Label { Text = message ?? "", Left = 22, Top = 52, Width = W - 44, Height = H - 52 - 54, ForeColor = Ink });
                var ok = Accent("OK", W - 22 - 96, H - 46, 96, 32);
                ok.DialogResult = DialogResult.OK;
                f.Controls.Add(ok);
                f.AcceptButton = ok; f.CancelButton = ok;
                f.ShowDialog();
            }
        }
    }

    // Move dialog — suggested folders, create-new-folder, and a searchable list. Opens instantly;
    // suggestions fill in async.
    // Resizable summary dialog: pick the language (re-summarizes on change), and a Reply button to
    // jump straight into composing a reply to the same email.
    internal class SummaryDialog : Form
    {
        private readonly TextBox _box;
        private readonly ComboBox _lang;
        private readonly Func<string, string> _summarize;   // language -> summary text
        public bool WantsReply { get; private set; }

        public SummaryDialog(string subject, Func<string, string> summarize)
        {
            _summarize = summarize;
            Text = "Axon intelligence — Summary";
            StartPosition = FormStartPosition.CenterScreen;
            ClientSize = new System.Drawing.Size(540, 460);
            MinimumSize = new System.Drawing.Size(440, 320);
            MaximizeBox = true; MinimizeBox = false; ShowIcon = false;
            BackColor = System.Drawing.Color.White;

            var langLbl = new Label { Left = 16, Top = 16, Width = 80, Height = 24, Text = "Summary in:",
                TextAlign = System.Drawing.ContentAlignment.MiddleLeft, Anchor = AnchorStyles.Top | AnchorStyles.Left };
            _lang = new ComboBox { Left = 100, Top = 13, Width = 200, Height = 24,
                DropDownStyle = ComboBoxStyle.DropDownList, Anchor = AnchorStyles.Top | AnchorStyles.Left };
            _lang.Items.AddRange(new object[] { "Auto (match the email)", "English", "Nederlands (Dutch)", "Français (French)" });
            _lang.SelectedIndex = 0;
            _lang.SelectedIndexChanged += (o, e) => Regenerate();
            _box = new TextBox { Left = 16, Top = 48, Width = ClientSize.Width - 32, Height = ClientSize.Height - 100,
                Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right,
                Multiline = true, ReadOnly = true, ScrollBars = ScrollBars.Vertical, BorderStyle = BorderStyle.FixedSingle,
                Text = "Summarizing…", Font = new System.Drawing.Font("Segoe UI", 9.75f), BackColor = System.Drawing.Color.White };
            var reply = new Button { Text = "Reply", Width = 90, Height = 28,
                Left = ClientSize.Width - 200, Top = ClientSize.Height - 40, Anchor = AnchorStyles.Bottom | AnchorStyles.Right };
            var close = new Button { Text = "Close", Width = 90, Height = 28, DialogResult = DialogResult.Cancel,
                Left = ClientSize.Width - 104, Top = ClientSize.Height - 40, Anchor = AnchorStyles.Bottom | AnchorStyles.Right };
            reply.Click += (o, e) => { WantsReply = true; DialogResult = DialogResult.OK; Close(); };
            Controls.AddRange(new Control[] { langLbl, _lang, _box, reply, close });
            CancelButton = close;
            Shown += (o, e) => Regenerate();
        }

        private static string LangCode(string display)
        {
            if (display.StartsWith("Eng")) return "English";
            if (display.StartsWith("Ned")) return "Dutch";
            if (display.StartsWith("Fr")) return "French";
            return "";   // Auto -> match the email
        }

        private void Regenerate()
        {
            string lang = LangCode(_lang.SelectedItem.ToString());
            SetText("Summarizing…");
            var th = new System.Threading.Thread(() =>
            {
                string s = _summarize(lang);
                SetText(string.IsNullOrEmpty(s) ? "Couldn't summarize (model unavailable)." : s);
            });
            th.IsBackground = true;
            th.Start();
        }

        private void SetText(string s)
        {
            if (IsDisposed) return;
            if (InvokeRequired) { try { BeginInvoke(new Action(() => SetText(s))); } catch { } return; }
            // Strip any stray markdown bold, and use CRLF so the TextBox shows line breaks.
            _box.Text = (s ?? "").Replace("**", "").Replace("\r\n", "\n").Replace("\n", "\r\n");
            _box.Select(0, 0);
        }
    }

    // Asks the user HOW to reply + in which language, then drafts it (off the UI thread).
    internal class ReplyPrompt : Form
    {
        private readonly TextBox _box;
        private readonly ComboBox _lang;
        private readonly Func<string, string, string> _drafter;   // (instruction, language) -> draft
        private readonly Button _draftBtn;
        private readonly Label _status;
        public string Draft { get; private set; }

        public ReplyPrompt(string subject, Func<string, string, string> drafter)
        {
            _drafter = drafter;
            Text = "Axon intelligence — Reply";
            StartPosition = FormStartPosition.CenterScreen;
            ClientSize = new System.Drawing.Size(500, 318);
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MaximizeBox = false; MinimizeBox = false; ShowIcon = false;
            BackColor = System.Drawing.Color.White;

            var t1 = new Label { Left = 20, Top = 18, Width = 460, Height = 20, Text = "Reply to: " + subject,
                Font = new System.Drawing.Font("Segoe UI", 10f, System.Drawing.FontStyle.Bold), AutoEllipsis = true };
            var t2 = new Label { Left = 20, Top = 48, Width = 460, Height = 20,
                Text = "Tell Axon how to reply (or leave blank for a professional reply):",
                ForeColor = System.Drawing.Color.FromArgb(110, 110, 120) };
            _box = new TextBox { Left = 20, Top = 74, Width = 460, Height = 110, Multiline = true,
                ScrollBars = ScrollBars.Vertical, BorderStyle = BorderStyle.FixedSingle,
                Font = new System.Drawing.Font("Segoe UI", 9.75f) };
            var langLbl = new Label { Left = 20, Top = 200, Width = 60, Height = 22, Text = "Reply in:",
                TextAlign = System.Drawing.ContentAlignment.MiddleLeft };
            _lang = new ComboBox { Left = 82, Top = 197, Width = 200, Height = 24, DropDownStyle = ComboBoxStyle.DropDownList };
            _lang.Items.AddRange(new object[] { "Auto (match the email)", "English", "Nederlands (Dutch)", "Français (French)" });
            _lang.SelectedIndex = 0;
            _status = new Label { Left = 20, Top = 240, Width = 260, Height = 20, Text = "",
                ForeColor = System.Drawing.Color.FromArgb(110, 110, 120) };
            _draftBtn = new Button { Text = "Draft reply", Left = 300, Top = 272, Width = 110, Height = 28 };
            var cancel = new Button { Text = "Cancel", Left = 415, Top = 272, Width = 65, Height = 28, DialogResult = DialogResult.Cancel };
            _draftBtn.Click += (o, e) => DoDraft();
            Controls.AddRange(new Control[] { t1, t2, _box, langLbl, _lang, _status, _draftBtn, cancel });
            AcceptButton = _draftBtn;
            CancelButton = cancel;
        }

        private static string LangCode(string display)
        {
            if (display.StartsWith("Eng")) return "English";
            if (display.StartsWith("Ned")) return "Dutch";
            if (display.StartsWith("Fr")) return "French";
            return "";   // Auto -> match the email
        }

        private void DoDraft()
        {
            _draftBtn.Enabled = false;
            _status.Text = "Drafting…";
            string instr = _box.Text;
            string lang = LangCode(_lang.SelectedItem.ToString());
            var th = new System.Threading.Thread(() =>
            {
                string d = _drafter(instr, lang);
                try { BeginInvoke(new Action(() => { Draft = d; DialogResult = DialogResult.OK; Close(); })); } catch { }
            });
            th.IsBackground = true;
            th.Start();
        }
    }

    // Asks the user WHAT the email should be about (+ language), then Axon writes it (off the UI thread).
    internal class WritePrompt : Form
    {
        private readonly TextBox _box;
        private readonly ComboBox _lang;
        private readonly Func<string, string, string> _writer;   // (description, language) -> email text
        private readonly Button _writeBtn;
        private readonly Label _status;
        public string Draft { get; private set; }

        public WritePrompt(Func<string, string, string> writer)
        {
            _writer = writer;
            Text = "Axon intelligence — Write email";
            StartPosition = FormStartPosition.CenterScreen;
            ClientSize = new System.Drawing.Size(500, 322);
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MaximizeBox = false; MinimizeBox = false; ShowIcon = false;
            BackColor = System.Drawing.Color.White;

            var t1 = new Label { Left = 20, Top = 16, Width = 460, Height = 20, Text = "Write an email with Axon",
                Font = new System.Drawing.Font("Segoe UI", 10f, System.Drawing.FontStyle.Bold) };
            var t2 = new Label { Left = 20, Top = 44, Width = 460, Height = 36,
                Text = "What should this email be about? Who it's to, the purpose, and any key points:",
                ForeColor = System.Drawing.Color.FromArgb(110, 110, 120) };
            _box = new TextBox { Left = 20, Top = 84, Width = 460, Height = 120, Multiline = true,
                ScrollBars = ScrollBars.Vertical, BorderStyle = BorderStyle.FixedSingle,
                Font = new System.Drawing.Font("Segoe UI", 9.75f) };
            var langLbl = new Label { Left = 20, Top = 218, Width = 60, Height = 22, Text = "Write in:",
                TextAlign = System.Drawing.ContentAlignment.MiddleLeft };
            _lang = new ComboBox { Left = 82, Top = 215, Width = 210, Height = 24, DropDownStyle = ComboBoxStyle.DropDownList };
            _lang.Items.AddRange(new object[] { "Auto (match my description)", "English", "Nederlands (Dutch)", "Français (French)" });
            _lang.SelectedIndex = 0;
            _status = new Label { Left = 20, Top = 258, Width = 260, Height = 20, Text = "",
                ForeColor = System.Drawing.Color.FromArgb(110, 110, 120) };
            _writeBtn = new Button { Text = "Write email", Left = 300, Top = 286, Width = 110, Height = 28 };
            var cancel = new Button { Text = "Cancel", Left = 415, Top = 286, Width = 65, Height = 28, DialogResult = DialogResult.Cancel };
            _writeBtn.Click += (o, e) => DoWrite();
            Controls.AddRange(new Control[] { t1, t2, _box, langLbl, _lang, _status, _writeBtn, cancel });
            AcceptButton = _writeBtn;
            CancelButton = cancel;
        }

        private static string LangCode(string display)
        {
            if (display.StartsWith("Eng")) return "English";
            if (display.StartsWith("Ned")) return "Dutch";
            if (display.StartsWith("Fr")) return "French";
            return "";   // Auto
        }

        private void DoWrite()
        {
            if (_box.Text.Trim().Length == 0) { _status.Text = "Describe the email first."; return; }
            _writeBtn.Enabled = false;
            _status.Text = "Writing…";
            string instr = _box.Text;
            string lang = LangCode(_lang.SelectedItem.ToString());
            var th = new System.Threading.Thread(() =>
            {
                string d = _writer(instr, lang);
                try { BeginInvoke(new Action(() => { Draft = d; DialogResult = DialogResult.OK; Close(); })); } catch { }
            });
            th.IsBackground = true;
            th.Start();
        }
    }

    // A friendly "when?" picker: one-click preset buttons (the common case) plus a calendar +
    // time list for a custom time. Used by Send Later and Follow up.
    internal class WhenPrompt : Form
    {
        private static readonly string[] Presets =
            { "In 1 hour", "This evening (6 PM)", "Tomorrow morning (8 AM)", "In 3 days", "Next Monday (9 AM)" };
        private readonly DateTimePicker _date;
        private readonly ComboBox _time;
        private readonly CheckBox _remind;
        public DateTime When { get; private set; }
        public bool Remind { get; private set; }

        public WhenPrompt(string title, string prompt, bool showRemind = false)
        {
            Text = "Axon intelligence — " + title;
            StartPosition = FormStartPosition.CenterScreen;
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MaximizeBox = false; MinimizeBox = false; ShowIcon = false;
            BackColor = System.Drawing.Color.White;
            int W = 380;

            var lbl = new Label { Left = 20, Top = 16, Width = W - 40, Height = 22, Text = prompt,
                Font = new System.Drawing.Font("Segoe UI", 10f, System.Drawing.FontStyle.Bold) };
            Controls.Add(lbl);

            int y = 50;
            for (int i = 0; i < Presets.Length; i++)
            {
                int idx = i;
                var b = new Button { Left = 20, Top = y, Width = W - 40, Height = 32, Text = "   " + Presets[i],
                    TextAlign = System.Drawing.ContentAlignment.MiddleLeft, FlatStyle = FlatStyle.System,
                    Font = new System.Drawing.Font("Segoe UI", 9.75f), Cursor = Cursors.Hand };
                b.Click += (o, e) => { When = Resolve(idx); Remind = _remind != null && _remind.Checked; DialogResult = DialogResult.OK; Close(); };
                Controls.Add(b);
                y += 36;
            }

            y += 8;
            var orLbl = new Label { Left = 20, Top = y + 3, Width = 56, Height = 22, Text = "Or pick:",
                ForeColor = System.Drawing.Color.FromArgb(110, 110, 120),
                TextAlign = System.Drawing.ContentAlignment.MiddleLeft };
            DateTime soon = DateTime.Now.AddHours(1);
            _date = new DateTimePicker { Left = 80, Top = y, Width = 158, Format = DateTimePickerFormat.Short };
            _date.Value = soon.Date;   // default to TODAY — same-day sends are the common case
            _time = new ComboBox { Left = 246, Top = y, Width = 114, DropDownStyle = ComboBoxStyle.DropDown };
            for (int h = 7; h <= 19; h++) { _time.Items.Add(h.ToString("00") + ":00"); _time.Items.Add(h.ToString("00") + ":30"); }
            _time.Text = soon.ToString("HH:mm");   // ~1 hour from now, so it's already in the future
            Controls.AddRange(new Control[] { orLbl, _date, _time });
            y += 44;

            if (showRemind)
            {
                _remind = new CheckBox { Left = 20, Top = y, Width = W - 40, Height = 22,
                    Text = "Remind me ~10 min before it sends",
                    Font = new System.Drawing.Font("Segoe UI", 9.25f) };
                Controls.Add(_remind);
                y += 30;
            }

            var set = new Button { Text = "Set custom", Left = W - 20 - 95 - 85, Top = y, Width = 95, Height = 28 };
            var cancel = new Button { Text = "Cancel", Left = W - 20 - 80, Top = y, Width = 80, Height = 28, DialogResult = DialogResult.Cancel };
            set.Click += (o, e) =>
            {
                int hh = 9, mm = 0;
                var parts = (_time.Text ?? "09:00").Split(':');
                int.TryParse(parts.Length > 0 ? parts[0] : "9", out hh);
                if (parts.Length > 1) int.TryParse(parts[1], out mm);
                When = _date.Value.Date.AddHours(hh).AddMinutes(mm);
                Remind = _remind != null && _remind.Checked;
                DialogResult = DialogResult.OK; Close();
            };
            Controls.AddRange(new Control[] { set, cancel });
            ClientSize = new System.Drawing.Size(W, y + 44);
            CancelButton = cancel;
        }

        private static DateTime Resolve(int i)
        {
            DateTime n = DateTime.Now;
            switch (i)
            {
                case 0: return n.AddHours(1);
                case 1: { DateTime e = n.Date.AddHours(18); return e > n ? e : n.AddHours(1); }
                case 2: return n.Date.AddDays(1).AddHours(8);
                case 3: return n.Date.AddDays(3).AddHours(9);
                case 4: { int add = (8 - (int)n.DayOfWeek) % 7; if (add == 0) add = 7; return n.Date.AddDays(add).AddHours(9); }
                default: return n.AddHours(1);
            }
        }
    }

    // Asks the user WHEN to schedule (any timeframe, free text) + the duration, before Axon drafts the invite.
    internal class SchedulePrompt : Form
    {
        private static readonly int[] Durations = { 15, 30, 45, 60, 90, 120 };
        private readonly TextBox _when;
        private readonly ComboBox _dur;
        public string When { get; private set; }
        public int DurationMinutes { get; private set; }

        public SchedulePrompt(string subject)
        {
            Text = "Axon intelligence — Schedule";
            StartPosition = FormStartPosition.CenterScreen;
            ClientSize = new System.Drawing.Size(500, 250);
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MaximizeBox = false; MinimizeBox = false; ShowIcon = false;
            BackColor = System.Drawing.Color.White;

            var t1 = new Label { Left = 20, Top = 18, Width = 460, Height = 20, Text = "Meeting about: " + subject,
                Font = new System.Drawing.Font("Segoe UI", 10f, System.Drawing.FontStyle.Bold), AutoEllipsis = true };
            var t2 = new Label { Left = 20, Top = 48, Width = 460, Height = 36,
                Text = "When? e.g. \"next Tuesday 2pm\", \"week of 15 Sept\", \"in 3 months\" — or leave blank to suggest the soonest free time:",
                ForeColor = System.Drawing.Color.FromArgb(110, 110, 120) };
            _when = new TextBox { Left = 20, Top = 88, Width = 460, Height = 26, BorderStyle = BorderStyle.FixedSingle,
                Font = new System.Drawing.Font("Segoe UI", 9.75f) };
            var durLbl = new Label { Left = 20, Top = 132, Width = 64, Height = 22, Text = "Duration:",
                TextAlign = System.Drawing.ContentAlignment.MiddleLeft };
            _dur = new ComboBox { Left = 86, Top = 129, Width = 160, Height = 24, DropDownStyle = ComboBoxStyle.DropDownList };
            _dur.Items.AddRange(new object[] { "15 minutes", "30 minutes", "45 minutes", "1 hour", "1.5 hours", "2 hours" });
            _dur.SelectedIndex = 1;   // 30 minutes
            var create = new Button { Text = "Create invite", Left = 300, Top = 200, Width = 110, Height = 28 };
            var cancel = new Button { Text = "Cancel", Left = 415, Top = 200, Width = 65, Height = 28, DialogResult = DialogResult.Cancel };
            create.Click += (o, e) =>
            {
                When = _when.Text.Trim();
                int i = _dur.SelectedIndex;
                DurationMinutes = (i >= 0 && i < Durations.Length) ? Durations[i] : 30;
                DialogResult = DialogResult.OK; Close();
            };
            Controls.AddRange(new Control[] { t1, t2, _when, durLbl, _dur, create, cancel });
            AcceptButton = create;
            CancelButton = cancel;
        }
    }

    internal class FolderPicker : Form
    {
        public string Chosen;
        public string CreateFolder;
        private readonly ListBox _list;
        private readonly TextBox _search;
        private readonly TextBox _newFolderBox;
        private string[] _all;
        private readonly Panel _suggPanel;

        public FolderPicker(string subject, string[] folders, string winTitle = "Move email",
            string headTitle = "Move this email to a folder", string createVerb = "Create && Move",
            string primaryVerb = "Move", string cancelVerb = "Keep in Inbox")
        {
            _all = folders ?? new string[0];
            Text = "Axon intelligence — " + winTitle;
            ClientSize = new System.Drawing.Size(500, 648);
            StartPosition = FormStartPosition.CenterScreen;
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MaximizeBox = false; MinimizeBox = false; ShowInTaskbar = false;
            BackColor = System.Drawing.Color.White;
            Font = new System.Drawing.Font("Segoe UI", 9.5F);
            int W = ClientSize.Width, H = ClientSize.Height;

            Controls.Add(Ui.Title(headTitle, 22, 18, W - 44));
            Controls.Add(Ui.Sub(subject ?? "(no subject)", 22, 46, W - 44));

            _suggPanel = new Panel { Left = 20, Top = 74, Width = W - 40, Height = 176 };
            _suggPanel.Controls.Add(Ui.Hint("Finding best matches…", 2, 6, W - 60, 30));
            Controls.Add(_suggPanel);

            Controls.Add(Ui.Caption("OR CREATE A NEW FOLDER", 22, 256, 300));
            _newFolderBox = new TextBox { Left = 22, Top = 274, Width = W - 44 - 8 - 148, BorderStyle = BorderStyle.FixedSingle };
            Controls.Add(_newFolderBox);
            var createBtn = Ui.Accent(createVerb, W - 22 - 148, 272, 148, 28);
            createBtn.Click += (o, e) =>
            {
                string nm = _newFolderBox.Text.Trim();
                if (nm.Length == 0) { Ui.Notify("Type a name for the new folder.", "Axon intelligence"); return; }
                CreateFolder = nm; DialogResult = DialogResult.OK; Close();
            };
            Controls.Add(createBtn);

            Controls.Add(Ui.Caption("OR PICK AN EXISTING FOLDER", 22, 312, 320));
            _search = new TextBox { Left = 22, Top = 330, Width = W - 44, BorderStyle = BorderStyle.FixedSingle };
            _search.TextChanged += (o, e) => Filter();
            Controls.Add(_search);
            _list = new ListBox
            {
                Left = 22, Top = 360, Width = W - 44, Height = H - 360 - 66,
                BorderStyle = BorderStyle.FixedSingle, DrawMode = DrawMode.OwnerDrawFixed, ItemHeight = 46, IntegralHeight = false
            };
            _list.DrawItem += (o, e) => Ui.DrawFolderItem(_list, e);
            _list.DoubleClick += (o, e) => PickFromList();
            Controls.Add(_list);
            Filter();

            var keep = Ui.Subtle(cancelVerb, W - 22 - 130, H - 50, 130, 34);
            keep.DialogResult = DialogResult.Cancel;
            Controls.Add(keep);
            CancelButton = keep;
            var move = Ui.Accent(primaryVerb, keep.Left - 8 - 104, H - 50, 104, 34);
            move.Click += (o, e) => PickFromList();
            Controls.Add(move);
        }

        public void SetFolders(string[] folders)
        {
            if (IsDisposed || Disposing) return;
            if (InvokeRequired) { try { BeginInvoke(new Action(() => SetFolders(folders))); } catch { } return; }
            _all = folders ?? new string[0];
            Filter();
        }

        // Back-compat: the Move feature calls this without reasons.
        public void SetSuggestions(string[] matches, string newFolder) { SetSuggestions(matches, null, newFolder); }

        public void SetSuggestions(string[] matches, string[] reasons, string newFolder)
        {
            if (IsDisposed || Disposing) return;
            if (InvokeRequired) { try { BeginInvoke(new Action(() => SetSuggestions(matches, reasons, newFolder))); } catch { } return; }
            int w = _suggPanel.Width;
            _suggPanel.Controls.Clear();
            if (matches != null && matches.Length > 0)
            {
                _suggPanel.Controls.Add(Ui.Caption("SUGGESTED FOLDERS", 2, 0, 300));
                int y = 18;
                for (int i = 0; i < matches.Length; i++)
                {
                    string n = matches[i];
                    // Show the full relative path (not just the leaf) so it's clear exactly where it saves.
                    var b = Ui.RowBtn("→   " + n, 2, y, w - 4, 28);
                    b.Click += (o, e) => { Chosen = n; DialogResult = DialogResult.OK; Close(); };
                    _suggPanel.Controls.Add(b); y += 31;
                }
                // #2 Pre-select the top suggestion in the list so the primary Save button uses it by default.
                for (int i = 0; i < _list.Items.Count; i++)
                    if (string.Equals(_list.Items[i].ToString(), matches[0], StringComparison.OrdinalIgnoreCase))
                    { _list.SelectedIndex = i; break; }
            }
            else
            {
                _suggPanel.Controls.Add(Ui.Hint("No existing folder clearly fits — create one above, or pick from the list.", 2, 6, w - 6, 40));
            }
            if (!string.IsNullOrEmpty(newFolder) && _newFolderBox.Text.Trim().Length == 0)
                _newFolderBox.Text = newFolder;
        }

        private void PickFromList()
        {
            if (_list.SelectedItem != null) { Chosen = _list.SelectedItem.ToString(); DialogResult = DialogResult.OK; Close(); }
        }

        private void Filter()
        {
            string q = _search.Text.Trim().ToLowerInvariant();
            _list.BeginUpdate();
            _list.Items.Clear();
            foreach (var f in _all)
                if (q == "" || f.ToLowerInvariant().Contains(q)) _list.Items.Add(f);
            if (_list.Items.Count > 0) _list.SelectedIndex = 0;
            _list.EndUpdate();
        }
    }

    // Download dialog — pick (or add) a disk folder to save the email into (no AI suggestion needed).
    internal class DownloadPicker : Form
    {
        public string Chosen;
        public System.Collections.Generic.List<string> Folders;
        private readonly ListBox _list;
        private readonly Label _emptyHint;

        public DownloadPicker(string subject, System.Collections.Generic.List<string> folders)
        {
            Folders = folders ?? new System.Collections.Generic.List<string>();
            Text = "Axon intelligence — Download email";
            ClientSize = new System.Drawing.Size(500, 404);
            StartPosition = FormStartPosition.CenterScreen;
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MaximizeBox = false; MinimizeBox = false; ShowInTaskbar = false;
            BackColor = System.Drawing.Color.White;
            Font = new System.Drawing.Font("Segoe UI", 9.5F);
            int W = ClientSize.Width, H = ClientSize.Height;

            Controls.Add(Ui.Title("Save this email to a folder", 22, 18, W - 44));
            Controls.Add(Ui.Sub(subject ?? "(no subject)", 22, 46, W - 44));

            Controls.Add(Ui.Caption("YOUR SAVE FOLDERS", 22, 82, 300));
            _list = new ListBox
            {
                Left = 22, Top = 102, Width = W - 44, Height = H - 102 - 66,
                BorderStyle = BorderStyle.FixedSingle, DrawMode = DrawMode.OwnerDrawFixed, ItemHeight = 46, IntegralHeight = false
            };
            foreach (var f in Folders) _list.Items.Add(f);
            if (_list.Items.Count > 0) _list.SelectedIndex = 0;
            _list.DrawItem += (o, e) => Ui.DrawFolderItem(_list, e);
            _list.DoubleClick += (o, e) => PickFromList();
            Controls.Add(_list);

            _emptyHint = Ui.Hint("No save folders yet — click “+ Add folder” to choose where emails are saved.", 30, 116, W - 88, 36);
            _emptyHint.Visible = Folders.Count == 0;
            Controls.Add(_emptyHint);
            _emptyHint.BringToFront();

            var add = Ui.Subtle("+  Add folder", 22, H - 50, 128, 34);
            add.Click += (o, e) => AddFolder();
            Controls.Add(add);
            var cancel = Ui.Subtle("Cancel", W - 22 - 100, H - 50, 100, 34);
            cancel.DialogResult = DialogResult.Cancel;
            Controls.Add(cancel);
            CancelButton = cancel;
            var save = Ui.Accent("Save here", cancel.Left - 8 - 108, H - 50, 108, 34);
            save.Click += (o, e) => PickFromList();
            Controls.Add(save);
        }

        private void AddFolder()
        {
            using (var fb = new FolderBrowserDialog())
            {
                fb.Description = "Choose a folder to save emails into";
                if (fb.ShowDialog() == DialogResult.OK && !string.IsNullOrEmpty(fb.SelectedPath))
                {
                    if (!Folders.Contains(fb.SelectedPath)) { Folders.Add(fb.SelectedPath); _list.Items.Add(fb.SelectedPath); }
                    _list.SelectedItem = fb.SelectedPath;
                    _emptyHint.Visible = false;
                }
            }
        }

        private void PickFromList()
        {
            if (_list.SelectedItem != null) { Chosen = _list.SelectedItem.ToString(); DialogResult = DialogResult.OK; Close(); }
            else Ui.Notify("Pick a folder, or click “+ Add folder”.");
        }
    }
}
