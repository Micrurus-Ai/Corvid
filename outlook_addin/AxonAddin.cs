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
    public partial class Connect : IDTExtensibility2, IRibbonExtensibility
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
