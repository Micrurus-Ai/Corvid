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
            // Put an "Axon" group on the built-in mail tab, EARLY (right after the New group), so the
            // buttons stay visible on the Home tab and don't fall into the "..." overflow.
            string tab, after;
            if (RibbonID == "Microsoft.Outlook.Explorer") { tab = "TabMail"; after = "GroupMailNew"; }          // main window Home
            else if (RibbonID == "Microsoft.Outlook.Mail.Read") { tab = "TabReadMessage"; after = "GroupRespond"; } // open email window (after Reply/Forward)
            else return null;
            string group =
                "<group id='axonGroup' label='Axon intelligence' insertAfterMso='" + after + "'>" +
                "<button id='axonMove' label='Move' size='normal' getImage='GetMoveImage' onAction='OnFile'/>" +
                "<button id='axonDownload' label='Download' size='normal' getImage='GetDownloadImage' onAction='OnDownload'/>" +
                "</group>";
            string ribbon = "<ribbon><tabs><tab idMso='" + tab + "'>" + group + "</tab></tabs></ribbon>";
            // Right-click menu on an email — always available, never pushed into the ribbon overflow.
            string ctx = "";
            if (RibbonID == "Microsoft.Outlook.Explorer")
            {
                ctx = "<contextMenus><contextMenu idMso='ContextMenuMailItem'>" +
                      "<button id='axonCtxMove' label='Move with Axon' getImage='GetMoveImage' onAction='OnFile'/>" +
                      "<button id='axonCtxDownload' label='Download with Axon' getImage='GetDownloadImage' onAction='OnDownload'/>" +
                      "</contextMenu></contextMenus>";
            }
            return "<customUI xmlns='http://schemas.microsoft.com/office/2009/07/customui'>" + ribbon + ctx + "</customUI>";
        }

        // --- custom ribbon image (the Axon-branded Move icon, distinct from Outlook's built-ins) ---
        private System.Drawing.Image _moveIcon, _downloadIcon;

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

        // Ask Axon's Python (agent.suggest_filing) for matching folders + a suggested new folder.
        private Filing SuggestFiling(string subject, string sender, string body, string[] folders)
        {
            var result = new Filing();
            try
            {
                string dir = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);   // ...\addin (or dev outlook_addin)
                string root = Path.GetFullPath(Path.Combine(dir, ".."));                         // Axon root
                var js = new System.Web.Script.Serialization.JavaScriptSerializer();
                string tmp = Path.Combine(Path.GetTempPath(), "axon_rank_" + Guid.NewGuid().ToString("N") + ".json");
                var payload = new System.Collections.Generic.Dictionary<string, object>
                { { "subject", subject }, { "sender", sender }, { "body", body }, { "folders", folders } };
                File.WriteAllText(tmp, js.Serialize(payload), new System.Text.UTF8Encoding(false));

                // Prefer the packaged app (installed); fall back to the dev venv + agent.py.
                System.Diagnostics.ProcessStartInfo psi;
                string exe = Path.Combine(root, "AxonIntelligence.exe");
                if (File.Exists(exe))
                {
                    psi = new System.Diagnostics.ProcessStartInfo(exe) { Arguments = "--suggest \"" + tmp + "\"" };
                }
                else
                {
                    string py = Path.Combine(root, "assistant", ".venv", "Scripts", "python.exe");
                    if (!File.Exists(py)) { try { File.Delete(tmp); } catch { } return result; }
                    psi = new System.Diagnostics.ProcessStartInfo(py)
                    {
                        Arguments = "-c \"import sys,json,agent; d=json.load(open(sys.argv[1],encoding='utf-8')); print(json.dumps(agent.suggest_filing(d['subject'],d['sender'],d['body'],d['folders'])))\" \"" + tmp + "\"",
                        WorkingDirectory = Path.Combine(root, "assistant")
                    };
                }
                psi.UseShellExecute = false;
                psi.RedirectStandardOutput = true;
                psi.CreateNoWindow = true;
                string outp;
                using (var p = System.Diagnostics.Process.Start(psi)) { outp = p.StandardOutput.ReadToEnd(); p.WaitForExit(25000); }
                try { File.Delete(tmp); } catch { }
                int b = outp.IndexOf('{'); int e = outp.LastIndexOf('}');
                if (b >= 0 && e > b)
                {
                    var d = (System.Collections.Generic.Dictionary<string, object>)js.DeserializeObject(outp.Substring(b, e - b + 1));
                    if (d.ContainsKey("matches") && d["matches"] is object[])
                    {
                        var arr = (object[])d["matches"];
                        var list = new System.Collections.Generic.List<string>();
                        foreach (var x in arr) if (x != null) list.Add(x.ToString());
                        result.matches = list.ToArray();
                    }
                    if (d.ContainsKey("new_folder") && d["new_folder"] != null) result.newFolder = d["new_folder"].ToString();
                }
            }
            catch { }
            return result;
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

        public void OnDownload(object control)
        {
            try
            {
                object m = GetSelectedMail();
                if (m == null) { Ui.Notify("Select an email first."); return; }
                dynamic mail = m;
                var folders = LoadDownloadFolders();
                string subject = ""; try { subject = (string)mail.Subject; } catch { }
                var dlg = new DownloadPicker(subject, folders);
                try
                {
                    var r = dlg.ShowDialog();
                    SaveDownloadFolders(dlg.Folders);            // persist any folders the user added
                    if (r == DialogResult.OK && !string.IsNullOrEmpty(dlg.Chosen))
                        SaveEmail(mail, dlg.Chosen, subject);
                }
                finally { dlg.Dispose(); }
            }
            catch (Exception ex) { Ui.Notify("Axon error: " + ex.Message, "Axon intelligence"); }
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
    internal class FolderPicker : Form
    {
        public string Chosen;
        public string CreateFolder;
        private readonly ListBox _list;
        private readonly TextBox _search;
        private readonly TextBox _newFolderBox;
        private readonly string[] _all;
        private readonly Panel _suggPanel;

        public FolderPicker(string subject, string[] folders)
        {
            _all = folders ?? new string[0];
            Text = "Axon intelligence — Move email";
            ClientSize = new System.Drawing.Size(500, 588);
            StartPosition = FormStartPosition.CenterScreen;
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MaximizeBox = false; MinimizeBox = false; ShowInTaskbar = false;
            BackColor = System.Drawing.Color.White;
            Font = new System.Drawing.Font("Segoe UI", 9.5F);
            int W = ClientSize.Width, H = ClientSize.Height;

            Controls.Add(Ui.Title("Move this email to a folder", 22, 18, W - 44));
            Controls.Add(Ui.Sub(subject ?? "(no subject)", 22, 46, W - 44));

            _suggPanel = new Panel { Left = 20, Top = 74, Width = W - 40, Height = 116 };
            _suggPanel.Controls.Add(Ui.Hint("Finding best matches…", 2, 6, W - 60, 30));
            Controls.Add(_suggPanel);

            Controls.Add(Ui.Caption("OR CREATE A NEW FOLDER", 22, 196, 300));
            _newFolderBox = new TextBox { Left = 22, Top = 214, Width = W - 44 - 8 - 148, BorderStyle = BorderStyle.FixedSingle };
            Controls.Add(_newFolderBox);
            var createBtn = Ui.Accent("Create && Move", W - 22 - 148, 212, 148, 28);
            createBtn.Click += (o, e) =>
            {
                string nm = _newFolderBox.Text.Trim();
                if (nm.Length == 0) { Ui.Notify("Type a name for the new folder.", "Axon intelligence"); return; }
                CreateFolder = nm; DialogResult = DialogResult.OK; Close();
            };
            Controls.Add(createBtn);

            Controls.Add(Ui.Caption("OR PICK AN EXISTING FOLDER", 22, 252, 320));
            _search = new TextBox { Left = 22, Top = 270, Width = W - 44, BorderStyle = BorderStyle.FixedSingle };
            _search.TextChanged += (o, e) => Filter();
            Controls.Add(_search);
            _list = new ListBox
            {
                Left = 22, Top = 300, Width = W - 44, Height = H - 300 - 66,
                BorderStyle = BorderStyle.FixedSingle, DrawMode = DrawMode.OwnerDrawFixed, ItemHeight = 46, IntegralHeight = false
            };
            _list.DrawItem += (o, e) => Ui.DrawFolderItem(_list, e);
            _list.DoubleClick += (o, e) => PickFromList();
            Controls.Add(_list);
            Filter();

            var keep = Ui.Subtle("Keep in Inbox", W - 22 - 130, H - 50, 130, 34);
            keep.DialogResult = DialogResult.Cancel;
            Controls.Add(keep);
            CancelButton = keep;
            var move = Ui.Accent("Move", keep.Left - 8 - 104, H - 50, 104, 34);
            move.Click += (o, e) => PickFromList();
            Controls.Add(move);
        }

        public void SetSuggestions(string[] matches, string newFolder)
        {
            if (IsDisposed || Disposing) return;
            if (InvokeRequired) { try { BeginInvoke(new Action(() => SetSuggestions(matches, newFolder))); } catch { } return; }
            int w = _suggPanel.Width;
            _suggPanel.Controls.Clear();
            if (matches != null && matches.Length > 0)
            {
                _suggPanel.Controls.Add(Ui.Caption("SUGGESTED FOLDERS", 2, 0, 300));
                int y = 18;
                foreach (var name in matches)
                {
                    string n = name;
                    var b = Ui.RowBtn("→   " + Ui.Leaf(name), 2, y, w - 4, 28);
                    b.Click += (o, e) => { Chosen = n; DialogResult = DialogResult.OK; Close(); };
                    _suggPanel.Controls.Add(b); y += 31;
                }
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
