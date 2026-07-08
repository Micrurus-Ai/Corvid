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
        public void OnStartupComplete(ref Array custom) { StartReminderService(); }
        public void OnBeginShutdown(ref Array custom) { try { if (_reminderTimer != null) _reminderTimer.Stop(); } catch { } }

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
                   "<menuSeparator id='axonSep2_" + idMso + "'/>" +
                   "<button id='axonSettings_" + idMso + "' label='Axon Settings' getImage='GetSettingsImage' onAction='OnSettings'/>" +
                   "</contextMenu>";
        }


        // --- custom ribbon image (the Axon-branded Move icon, distinct from Outlook's built-ins) ---
        private System.Drawing.Image _moveIcon, _downloadIcon, _summarizeIcon, _replyIcon, _scheduleIcon,
                                     _followUpIcon, _sendLaterIcon, _writeIcon, _attachIcon, _settingsIcon;

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

        public stdole.IPictureDisp GetSettingsImage(object control)
        {
            try { return RibbonImage.Get(LoadIcon("axon-settings.png", ref _settingsIcon)); } catch { return null; }
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
            // Never leave the 'create new folder' box empty when nothing matched — propose a name.
            if ((result.matches == null || result.matches.Length == 0) && string.IsNullOrWhiteSpace(result.newFolder))
                result.newFolder = FallbackFolderName(subject, sender);
            return result;
        }

        // Derive a sensible new-folder name from the subject (or sender) when the model doesn't give one.
        private static string FallbackFolderName(string subject, string sender)
        {
            string s = (subject ?? "").Trim();
            s = System.Text.RegularExpressions.Regex.Replace(s, @"^((RE|FW|FWD|AW|VS|TR)\s*:\s*)+", "",
                                                             System.Text.RegularExpressions.RegexOptions.IgnoreCase).Trim();
            foreach (var sep in new[] { " | ", " - ", " – ", " — ", ": " })
            { int i = s.IndexOf(sep); if (i > 2) { s = s.Substring(0, i).Trim(); break; } }
            var words = s.Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);
            if (words.Length > 4) s = string.Join(" ", words, 0, 4);
            if (string.IsNullOrWhiteSpace(s)) s = (sender ?? "").Trim();
            foreach (var c in Path.GetInvalidFileNameChars()) s = s.Replace(c.ToString(), "");
            s = s.Trim();
            if (s.Length > 40) s = s.Substring(0, 40).Trim();
            return string.IsNullOrWhiteSpace(s) ? "New folder" : s;
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
            if (b.Length > 4000) b = b.Substring(0, 4000);
            string prompt =
                "You are filing an email into one of the user's Outlook folders.\n\n" +
                "Email\n  Subject: " + subject + "\n  From: " + sender + "\n  Body:\n" + b + "\n\n" +
                "The user's folders (name, with full path for nested ones):\n" + js.Serialize(folders) + "\n\n" +
                "READ THE BODY carefully to understand what the email is really about — the subject alone is not " +
                "enough. Decide based on the email's actual TOPIC and WHO it is from. Pick the folder whose name/path " +
                "most closely matches that topic; the sender's company/domain is a strong hint. List the " +
                "best-fitting folders first (up to 5), but ONLY include a folder that is a genuinely good fit " +
                "(do not pad the list). If none clearly fit, leave matches empty and you MUST propose a short " +
                "new_folder name based on the email's topic or sender — never leave both matches and new_folder empty.\n" +
                "Reply with ONLY JSON: {\"matches\": [up to 5 folder names copied EXACTLY from the list, best " +
                "first], \"new_folder\": \"a short (1-3 word) new folder name; REQUIRED whenever matches is empty\"}";
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
                               category = Field(info, "category"), year = Field(info, "year"), sap = Field(info, "sap");
                        if (string.IsNullOrWhiteSpace(year)) year = DateTime.Now.Year.ToString();
                        string sender = ""; try { sender = (string)mail.SenderName; } catch { }
                        string body = ""; try { body = (string)mail.Body; } catch { }
                        string baseDir = ResolveBaseDir(cfg, category, code);
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
                            string root = _archiveBaseDir ?? cfg.PathFor("");
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
            public string SaveMode = "both", Subfolder = "";
            // Named base folders the user defined (label -> path), in order. Not hardcoded to
            // client/supplier — the user can add any labels, and as many as they want.
            public System.Collections.Generic.List<System.Collections.Generic.KeyValuePair<string, string>> Bases =
                new System.Collections.Generic.List<System.Collections.Generic.KeyValuePair<string, string>>();
            public System.Collections.Generic.Dictionary<string, string> Codes =
                new System.Collections.Generic.Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            public bool Ready { get { foreach (var b in Bases) if (!string.IsNullOrWhiteSpace(b.Value)) return true; return false; } }
            public string[] Labels
            {
                get { var l = new System.Collections.Generic.List<string>(); foreach (var b in Bases) if (!string.IsNullOrWhiteSpace(b.Value)) l.Add(b.Key); return l.ToArray(); }
            }
            public string PathFor(string label)
            {
                foreach (var b in Bases) if (string.Equals(b.Key, label, StringComparison.OrdinalIgnoreCase) && !string.IsNullOrWhiteSpace(b.Value)) return b.Value;
                foreach (var b in Bases) if (!string.IsNullOrWhiteSpace(b.Value)) return b.Value;   // fallback: first defined
                return "";
            }
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
                var basesArr = d.ContainsKey("bases") ? d["bases"] as object[] : null;
                if (basesArr != null)
                    foreach (var o in basesArr)
                    {
                        var bd = o as System.Collections.Generic.Dictionary<string, object>;
                        if (bd == null) continue;
                        string nm = bd.ContainsKey("name") && bd["name"] != null ? bd["name"].ToString().Trim() : "";
                        string pth = bd.ContainsKey("path") && bd["path"] != null ? bd["path"].ToString().Trim() : "";
                        if (!string.IsNullOrWhiteSpace(nm)) cfg.Bases.Add(new System.Collections.Generic.KeyValuePair<string, string>(nm, pth));
                    }
                if (cfg.Bases.Count == 0)   // backward-compat: old fixed client_base / supplier_base
                {
                    string cb = d.ContainsKey("client_base") && d["client_base"] != null ? d["client_base"].ToString().Trim() : "";
                    string sb = d.ContainsKey("supplier_base") && d["supplier_base"] != null ? d["supplier_base"].ToString().Trim() : "";
                    if (!string.IsNullOrWhiteSpace(cb)) cfg.Bases.Add(new System.Collections.Generic.KeyValuePair<string, string>("Clients", cb));
                    if (!string.IsNullOrWhiteSpace(sb)) cfg.Bases.Add(new System.Collections.Generic.KeyValuePair<string, string>("Suppliers", sb));
                }
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
            if (body.Length > 4000) body = body.Substring(0, 4000);
            var js = new System.Web.Script.Serialization.JavaScriptSerializer();
            string mapJson = js.Serialize(cfg.Codes);
            string labels = js.Serialize(cfg.Labels);
            string prompt =
                "Read this email and reply with ONLY JSON: {\"category\":\"\",\"code\":\"\",\"company\":\"\"," +
                "\"year\":\"\",\"sap\":\"\"}. \"category\" = which ONE of the user's archive folders this email belongs in " +
                "(choose exactly one label from this list, or empty if none clearly fits): " + labels + ". " +
                "Determine the sender's COUNTRY from the email domain, any phone numbers " +
                "(+32 Belgium, +49 Germany, +31 Netherlands, +33 France, etc.), and address, then set \"code\" using this " +
                "Country->code map: " + mapJson + " (empty if the country isn't in the map). \"company\" = the sender's " +
                "company name. \"year\" = a 4-digit year from the email, else the current year. \"sap\" = the order/" +
                "SAP number in the subject if there is one, else empty.\n\nFrom: " + sender + " <" + senderEmail + ">\n" +
                "Subject: " + subject + "\n\n" + body;
            string text = ModelComplete(prompt, 0);
            var m = System.Text.RegularExpressions.Regex.Match(text ?? "", "\\{[\\s\\S]*\\}");
            if (!m.Success) return null;
            try { return js.DeserializeObject(m.Value) as System.Collections.Generic.Dictionary<string, object>; }
            catch { return null; }
        }

        private string ResolveBaseDir(ArchiveCfg cfg, string category, string code)
        {
            string bas = cfg.PathFor(category);
            bas = (bas ?? "").TrimEnd('\\', '/');
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

                string b = body ?? ""; if (b.Length > 3000) b = b.Substring(0, 3000);
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
            return "Summarize this email THOROUGHLY in PLAIN TEXT only — no markdown, no asterisks, no '#'. " +
                "Capture all the important information: who is involved, what they are asking or saying, and every " +
                "specific detail that matters (names, dates, amounts, quantities, order/reference numbers, deadlines, " +
                "links, attachments, and any questions raised). Do NOT over-compress — it is better to keep a detail " +
                "than to drop it. Use this layout, with a blank line between sections:\n\n" +
                "Gist: <2-3 sentences giving the full picture>\n\n" +
                "Key points:\n- <point>\n- <point>\n- <point>\n(list every meaningful point, one per line — include the specifics)\n\n" +
                "Action: <what the reader should do, including any deadline or detail, or 'None'>\n\n" +
                langLine + "\n\nSubject: " + subject + "\n\n" + body;
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

        // --- Reminder service --------------------------------------------------------------------
        // Fires Follow-up / Send-Later reminders from INSIDE Outlook, so they work even without the
        // floating dot. Reads the same JSON the add-in writes and uses the SAME 'seen' files as the
        // dot (keys match), and it stands down if the dot is running — so the two never double-fire.
        private System.Windows.Forms.Timer _reminderTimer;

        private void StartReminderService()
        {
            try
            {
                _reminderTimer = new System.Windows.Forms.Timer();
                _reminderTimer.Interval = 60000;                 // check every minute
                _reminderTimer.Tick += (o, e) => CheckRemindersSafe();
                _reminderTimer.Start();
                var first = new System.Windows.Forms.Timer();    // and once, ~10s after Outlook opens
                first.Interval = 10000;
                first.Tick += (o, e) => { first.Stop(); first.Dispose(); CheckRemindersSafe(); };
                first.Start();
            }
            catch { }
        }

        private void CheckRemindersSafe() { try { CheckReminders(); } catch { } }

        private static bool DotIsRunning()
        {
            try { return System.Diagnostics.Process.GetProcessesByName("AxonIntelligence").Length > 0; }
            catch { return false; }
        }

        private static string AxonDir()
        {
            return Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "AxonOutlook");
        }

        private static System.Collections.Generic.List<System.Collections.Generic.Dictionary<string, object>> LoadArray(string path)
        {
            var outList = new System.Collections.Generic.List<System.Collections.Generic.Dictionary<string, object>>();
            try
            {
                if (!File.Exists(path)) return outList;
                var js = new System.Web.Script.Serialization.JavaScriptSerializer();
                var arr = js.DeserializeObject(File.ReadAllText(path)) as object[];
                if (arr != null)
                    foreach (var o in arr)
                    {
                        var d = o as System.Collections.Generic.Dictionary<string, object>;
                        if (d != null) outList.Add(d);
                    }
            }
            catch { }
            return outList;
        }

        private static System.Collections.Generic.HashSet<string> LoadSeen(string path)
        {
            var set = new System.Collections.Generic.HashSet<string>();
            try
            {
                if (!File.Exists(path)) return set;
                var js = new System.Web.Script.Serialization.JavaScriptSerializer();
                var arr = js.DeserializeObject(File.ReadAllText(path)) as object[];
                if (arr != null) foreach (var o in arr) if (o != null) set.Add(o.ToString());
            }
            catch { }
            return set;
        }

        private static void SaveSeen(string path, System.Collections.Generic.HashSet<string> set)
        {
            try
            {
                var js = new System.Web.Script.Serialization.JavaScriptSerializer();
                Directory.CreateDirectory(Path.GetDirectoryName(path));
                File.WriteAllText(path, js.Serialize(new System.Collections.Generic.List<string>(set)));
            }
            catch { }
        }

        private static string F(System.Collections.Generic.Dictionary<string, object> d, string k)
        {
            object v; return (d != null && d.TryGetValue(k, out v) && v != null) ? v.ToString() : "";
        }

        private void CheckReminders()
        {
            if (DotIsRunning()) return;   // the dot owns reminders when it's running; avoid duplicates
            DateTime now = DateTime.Now;
            string dir = AxonDir();

            // Follow-ups (seen key: id|due, matching the dot).
            var fItems = LoadArray(Path.Combine(dir, "followups.json"));
            if (fItems.Count > 0)
            {
                string seenPath = Path.Combine(dir, "followups_seen.json");
                var seen = LoadSeen(seenPath);
                bool changed = false;
                foreach (var it in fItems)
                {
                    string due = F(it, "due");
                    string key = F(it, "id") + "|" + due;
                    if (seen.Contains(key)) continue;
                    DateTime d;
                    if (!DateTime.TryParse(due, out d) || d > now) continue;
                    DateTime created; DateTime.TryParse(F(it, "created"), out created);
                    if (HasReply(F(it, "id"), created)) { seen.Add(key); changed = true; continue; }
                    string who = F(it, "who"), subj = F(it, "subject");
                    ShowReminder("Follow up" + (string.IsNullOrEmpty(who) ? "" : " with " + who),
                                 string.IsNullOrEmpty(subj) ? "(no subject)" : subj);
                    seen.Add(key); changed = true;
                }
                if (changed) SaveSeen(seenPath, seen);
            }

            // Scheduled sends (seen keys: to|when|remind and to|when|sent, matching the dot).
            var sItems = LoadArray(Path.Combine(dir, "scheduled.json"));
            if (sItems.Count > 0)
            {
                string seenPath = Path.Combine(dir, "scheduled_seen.json");
                var seen = LoadSeen(seenPath);
                bool changed = false;
                foreach (var it in sItems)
                {
                    string bas = F(it, "to") + "|" + F(it, "when");
                    string subj = F(it, "subject");
                    string remindAt = F(it, "remind_at");
                    if (!string.IsNullOrEmpty(remindAt))
                    {
                        string rk = bas + "|remind"; DateTime rt;
                        if (!seen.Contains(rk) && DateTime.TryParse(remindAt, out rt) && rt <= now)
                        {
                            ShowReminder("Axon will send this scheduled email at " + F(it, "when") + ".",
                                string.IsNullOrEmpty(subj) ? "Scheduled email" : subj);
                            seen.Add(rk); changed = true;
                        }
                    }
                    string sk = bas + "|sent"; DateTime w;
                    if (!seen.Contains(sk) && DateTime.TryParse(F(it, "when"), out w) && w <= now)
                    {
                        ShowReminder("This scheduled email has now been sent from your Outbox.",
                            string.IsNullOrEmpty(subj) ? "Scheduled email" : subj);
                        seen.Add(sk); changed = true;
                    }
                }
                if (changed) SaveSeen(seenPath, seen);
            }
        }

        // Best-effort: did a message in this conversation arrive after the follow-up was set?
        private bool HasReply(string entryId, DateTime created)
        {
            if (string.IsNullOrEmpty(entryId)) return false;
            try
            {
                dynamic app = _app; if (app == null) return false;
                dynamic ns = app.GetNamespace("MAPI");
                dynamic item = ns.GetItemFromID(entryId);
                if (item == null) return false;
                dynamic conv = null; try { conv = item.GetConversation(); } catch { }
                if (conv == null) return false;
                dynamic table = conv.GetTable();
                try { table.Columns.Add("ReceivedTime"); } catch { }
                while (!(bool)table.EndOfTable)
                {
                    dynamic row = table.GetNextRow();
                    try
                    {
                        object rto = row["ReceivedTime"];
                        if (rto != null && Convert.ToDateTime(rto) > created.AddMinutes(1)) return true;
                    }
                    catch { }
                }
            }
            catch { }
            return false;
        }

        // A reliable reminder styled like Outlook's own "1 Reminder" window: light, item in bold,
        // Dismiss + Snooze. Non-modal + topmost so it's dependable (Outlook's native one is flaky).
        private void ShowReminder(string kind, string item)
        {
            try
            {
                var f = new Form
                {
                    Text = "1 Reminder",
                    FormBorderStyle = FormBorderStyle.FixedDialog,
                    MaximizeBox = false, MinimizeBox = false,
                    ShowInTaskbar = false, TopMost = true,
                    BackColor = System.Drawing.Color.White,
                    Font = new System.Drawing.Font("Segoe UI", 9F),
                    ClientSize = new System.Drawing.Size(410, 152),
                    StartPosition = FormStartPosition.CenterScreen
                };
                int W = f.ClientSize.Width;
                f.Controls.Add(new Label
                {
                    Text = "🔔", Left = 16, Top = 18, Width = 34, Height = 34,   // 🔔
                    Font = new System.Drawing.Font("Segoe UI Emoji", 15F),
                    ForeColor = System.Drawing.Color.FromArgb(198, 118, 28)
                });
                f.Controls.Add(new Label
                {
                    Text = string.IsNullOrEmpty(item) ? "(no subject)" : item,
                    Left = 58, Top = 18, Width = W - 74, Height = 24, AutoEllipsis = true,
                    Font = new System.Drawing.Font("Segoe UI", 10.5F, System.Drawing.FontStyle.Bold),
                    ForeColor = System.Drawing.Color.FromArgb(28, 28, 28)
                });
                f.Controls.Add(new Label
                {
                    Text = kind, Left = 58, Top = 46, Width = W - 74, Height = 34, AutoEllipsis = true,
                    ForeColor = System.Drawing.Color.FromArgb(96, 96, 96)
                });
                f.Controls.Add(new Label { Left = 0, Top = 100, Width = W, Height = 1, BackColor = System.Drawing.Color.FromArgb(224, 224, 224) });

                var dismiss = new Button { Text = "Dismiss", Left = W - 16 - 96, Top = 112, Width = 96, Height = 30, FlatStyle = FlatStyle.System };
                dismiss.Click += (o, e) => { try { f.Close(); } catch { } };
                var snooze = new Button { Text = "Snooze 5 min", Left = W - 16 - 96 - 8 - 116, Top = 112, Width = 116, Height = 30, FlatStyle = FlatStyle.System };
                snooze.Click += (o, e) =>
                {
                    try { f.Close(); } catch { }
                    var re = new System.Windows.Forms.Timer { Interval = 5 * 60 * 1000 };
                    re.Tick += (o2, e2) => { re.Stop(); re.Dispose(); ShowReminder(kind, item); };
                    re.Start();
                };
                f.Controls.Add(dismiss); f.Controls.Add(snooze);
                f.AcceptButton = dismiss;
                f.FormClosed += (o, e) => { try { f.Dispose(); } catch { } };
                f.Show();
            }
            catch { }
        }

        // --- Settings (standalone add-in has no dot, so it needs its own settings UI) ------------
        public void OnSettings(object control)
        {
            try { using (var f = new SettingsForm(this)) { f.ShowDialog(); } }
            catch (Exception ex) { Ui.Notify("Axon error: " + ex.Message, "Axon intelligence"); }
        }

        // Read the user's recent Sent mail and derive a writing-style guide (used by Reply / Write).
        // Runs the same idea as the dot's 'learn tone', but entirely inside the add-in.
        public string LearnToneFromSent()
        {
            try
            {
                dynamic app = _app; if (app == null) return "";
                dynamic ns = app.GetNamespace("MAPI");
                dynamic sent = ns.GetDefaultFolder(5);   // olFolderSentMail
                dynamic items = sent.Items;
                try { items.Sort("[SentOn]", true); } catch { }
                var sb = new System.Text.StringBuilder();
                int n = 0;
                dynamic m = null; try { m = items.GetFirst(); } catch { }
                while (m != null && n < 25)
                {
                    string b = ""; try { b = (string)m.Body; } catch { }
                    foreach (var mark in new[] { "\r\nFrom:", "\r\nVan:", "\r\nSent:", "-----Original", "-----Oorspronkelijk" })
                    { int i = b.IndexOf(mark); if (i > 0) { b = b.Substring(0, i); } }
                    b = (b ?? "").Trim();
                    if (b.Length >= 20)
                    {
                        if (b.Length > 1000) b = b.Substring(0, 1000);
                        sb.AppendLine("=====EMAIL====="); sb.AppendLine(b); n++;
                    }
                    try { m = items.GetNext(); } catch { m = null; }
                }
                if (n == 0) return "";
                string prompt =
                    "Below are recent emails the USER has SENT (their own outgoing messages). Describe ONLY how the " +
                    "user writes; ignore any quoted text. Write a concise STYLE GUIDE another writer could follow: " +
                    "greeting, sign-off, formality/warmth, sentence length, whether they use bullets or emojis, and " +
                    "which languages they write in. Plain text with '- ' bullets, no Markdown, no asterisks, under " +
                    "200 words.\n\n" + sb.ToString();
                string tone = ModelComplete(prompt, 0.2);
                return (tone ?? "").Replace("*", "").Trim();
            }
            catch { return ""; }
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
}
