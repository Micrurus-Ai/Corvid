// Axon Outlook add-in — Email-archive Download: config, entity extraction, folder suggestions, saving.  (partial of Connect; split out of AxonAddin.cs.)
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
    }
}
