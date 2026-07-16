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
                        string senderEmail = ""; try { senderEmail = (string)mail.SenderEmailAddress; } catch { }
                        string body = ""; try { body = (string)mail.Body; } catch { }
                        string baseDir = ResolveBaseDir(cfg, category, code);
                        _archiveBaseDir = baseDir;
                        _archiveCompany = company;
                        System.Collections.Generic.List<string> matches, existing, reasons; string newRel;
                        BuildArchiveSuggestions(baseDir, subject, sender, senderEmail, body, year, company, sap, code,
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
            // Axon Group's standard country codes, seeded so Download works out of the box for all four
            // countries without anyone configuring them. A user's own Settings entry overrides these.
            foreach (var def in new[] { new[] { "Belgium", "AB" }, new[] { "Germany", "AD" },
                                        new[] { "France", "AF" }, new[] { "Netherlands", "AN" } })
                if (!cfg.Codes.ContainsKey(def[0])) cfg.Codes[def[0]] = def[1];
            // Default archive root, so Download works with NO setup at all. A user who maps the share to a
            // different drive/path can still override it in Settings > Archive folders.
            if (!cfg.Ready) cfg.Bases.Add(new System.Collections.Generic.KeyValuePair<string, string>("Sales", @"T:\IF\Sales"));
            return cfg;
        }

        private System.Collections.Generic.Dictionary<string, object> ExtractArchiveInfo(dynamic mail, ArchiveCfg cfg)
        {
            string subject = ""; try { subject = (string)mail.Subject; } catch { }
            string sender = ""; try { sender = (string)mail.SenderName; } catch { }
            string senderEmail = ""; try { senderEmail = (string)mail.SenderEmailAddress; } catch { }
            string body = ""; try { body = (string)mail.Body; } catch { }
            // Read the whole thread — order emails are usually several messages of back-and-forth, and the
            // customer/order details often sit in an earlier message, not the newest one.
            if (body.Length > 24000) body = body.Substring(0, 24000);
            var js = new System.Web.Script.Serialization.JavaScriptSerializer();
            string mapJson = js.Serialize(cfg.Codes);
            string labels = js.Serialize(cfg.Labels);
            string prompt =
                "Read this email THREAD and reply with ONLY JSON: {\"category\":\"\",\"code\":\"\",\"company\":\"\"," +
                "\"year\":\"\",\"sap\":\"\"}.\n" +
                "This may be a thread with several back-and-forth messages (newest at the top); use ALL of it. " +
                "It may also be FORWARDED (subject starts with FW/FWD/TR/Doorst, or the body quotes earlier " +
                "messages). Either way, the party this email is ABOUT is the ORIGINAL EXTERNAL client/company " +
                "in the conversation, NOT the colleague who forwarded it. Identify that external client and use " +
                "IT for company and country, reading the whole thread — quoted messages, signatures and any " +
                "'From:' lines — not just the visible sender.\n" +
                "\"category\" = which ONE of the user's archive folders this email belongs in (choose exactly one " +
                "label from this list, or empty if none clearly fits): " + labels + ".\n" +
                "\"company\" = the external client's company name.\n" +
                "\"code\" = that client's COUNTRY code. Determine the client's country from their email domain, " +
                "phone numbers (+32 Belgium, +49 Germany, +31 Netherlands, +33 France, etc.), address or signature, " +
                "then map it with this Country->code map: " + mapJson + " (empty if the country isn't in the map).\n" +
                "\"sap\" = the order or SAP number, usually the leading number in the subject (for a subject like " +
                "'FW: NNNNN - description' that would be NNNNN). Empty if there is none.\n" +
                "\"year\" = a 4-digit year from the email, else the current year.\n\n" +
                "Visible sender (may be an internal forwarder): " + sender + " <" + senderEmail + ">\n" +
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

        // Axon Group's Sales archive follows a fixed shape: base \ {code} \ SOP \ {year} \ {client} \
        // {order} \ {category}. So instead of searching the (slow, huge) share we NAVIGATE straight down
        // it — one immediate-child listing per level, ~10 ms each. SOP is the orders folder; the code
        // comes from the country map; year/client/order are read from the email and matched to the folder.
        // The per-order category folders (Documents, Order, Quotation, MC/MI/MS...) vary between orders, so
        // we scan just that ONE order (cheap) and let the user pick. The picker always lets them choose a
        // different folder, so a wrong guess is never a dead end.
        private const string OrdersTypeFolder = "SOP";

        private void BuildArchiveSuggestions(string baseDir, string subject, string sender, string senderEmail, string body,
            string year, string company, string sap, string code,
            out System.Collections.Generic.List<string> matches, out System.Collections.Generic.List<string> reasons,
            out string newRel, out System.Collections.Generic.List<string> existing)
        {
            var mm = new System.Collections.Generic.List<string>();
            var rr = new System.Collections.Generic.List<string>();
            matches = mm; reasons = rr;
            existing = new System.Collections.Generic.List<string>();
            newRel = "";

            Action<string, string> add = (rel, why) =>
            {
                if (string.IsNullOrEmpty(rel) || mm.Count >= 6) return;
                foreach (var m in mm) if (string.Equals(m, rel, StringComparison.OrdinalIgnoreCase)) return;
                mm.Add(rel); rr.Add(why);
            };

            try
            {
                if (!Directory.Exists(baseDir)) return;
                var sw = System.Diagnostics.Stopwatch.StartNew();
                const int budgetMs = 6000;

                // --- Deterministic descent: one listing per level, straight down the known template. ---
                string dir = baseDir;
                string codeDir   = FindChild(dir, code, true);        if (codeDir   != null) dir = codeDir;   // Sales\AB
                string sopDir    = FindChild(dir, OrdersTypeFolder, true); if (sopDir != null) dir = sopDir;   // ...\SOP
                string yearDir   = FindChild(dir, year, false);       if (yearDir   != null) dir = yearDir;   // ...\2026 (14334-)
                string clientDir = FindChild(dir, company, false);    if (clientDir != null) dir = clientDir; // ...\Voestalpine
                string orderDir  = FindChild(dir, sap, false);        // ...\14457_ADD low noise
                // Fallback: if the client wasn't identified (e.g. a forward), the order sits a level or two
                // below where we stopped — a small bounded search from here still finds it by its number.
                if (orderDir == null && !string.IsNullOrEmpty(sap))
                {
                    orderDir = FindDescendant(dir, sap, 2, sw, budgetMs);
                    if (orderDir != null && clientDir == null)
                        try { clientDir = System.IO.Directory.GetParent(orderDir).FullName; } catch { }
                }

                if (orderDir != null)
                {
                    // Found the order. Offer it and its own subfolders (the categories) for the user to pick.
                    AddRel(baseDir, orderDir, existing);
                    CollectSubtreeRel(baseDir, orderDir, 3, existing, sw, budgetMs);
                    // Also list the client's OTHER orders, so picking a different order is easy.
                    if (clientDir != null) CollectSubtreeRel(baseDir, clientDir, 1, existing, sw, budgetMs);

                    string orderRel = RelOf(baseDir, orderDir);

                    // Auto-pick the category from Axon Group's filing rules — Order vs Quotation, and
                    // MC = customer / MI = internal group company / MS = supplier. Pre-selected as the top
                    // suggestion; the user can still choose a different one.
                    string catRel, catReason;
                    PickOrderCategory(subject, sender, senderEmail, body, company, orderRel, existing, out catRel, out catReason);
                    if (!string.IsNullOrEmpty(catRel)) add(catRel, catReason);

                    // Then the order folder itself, then its other category subfolders.
                    add(orderRel, "this order");
                    foreach (var rel in existing)
                        if (rel.StartsWith(orderRel + "\\", StringComparison.OrdinalIgnoreCase)
                            && rel.Split('\\').Length == orderRel.Split('\\').Length + 1)
                            add(rel, "category");
                }
                else
                {
                    // Order folder not there yet. Show what IS at the deepest level we reached so the user
                    // can pick or navigate, and propose a new order folder that follows the template.
                    CollectSubtreeRel(baseDir, dir, 1, existing, sw, budgetMs);

                    string deepest; var tail = new System.Collections.Generic.List<string>();
                    if (clientDir != null)   { deepest = clientDir; tail.Add(sap); }
                    else if (yearDir != null){ deepest = yearDir;   tail.Add(company); tail.Add(sap); }
                    else if (sopDir != null) { deepest = sopDir;    tail.Add(year); tail.Add(company); tail.Add(sap); }
                    else if (codeDir != null){ deepest = codeDir;   tail.Add(OrdersTypeFolder); tail.Add(year); tail.Add(company); tail.Add(sap); }
                    else                     { deepest = baseDir;   tail.Add(code); tail.Add(OrdersTypeFolder); tail.Add(year); tail.Add(company); tail.Add(sap); }
                    var relParts = new System.Collections.Generic.List<string>();
                    string dr = RelOf(baseDir, deepest); if (!string.IsNullOrEmpty(dr)) relParts.Add(dr);
                    foreach (var t in tail) if (!string.IsNullOrWhiteSpace(t)) relParts.Add(t.Trim());
                    newRel = string.Join("\\", relParts);
                }
            }
            catch { }
        }

        // Axon Group's own email domains — the reliable way to tell MI (internal) from MS (supplier):
        // a correspondent whose email domain is one of these is INTERNAL. Extend as new group companies
        // are added.
        private const string InternalDomains =
            "almeco.be, axongroup.com, noviso.eu, coateq.be, proceq.eu, dimplesteel.com, " +
            "akwaplus.be, enviro-tech.nl, pcacontrol.com, pcawater.com, pca-air.com";

        // Inside an order the email is filed by TYPE (Order = an order, Quotation = a quote) and by WHO it
        // is with (MC = the customer/client, MI = an internal group company, MS = a supplier). Ask the
        // model to classify, then match that to a real subfolder of this order. Returns the folder to
        // pre-select and a one-word reason, or nulls if it can't decide / the folder isn't there.
        private void PickOrderCategory(string subject, string sender, string senderEmail, string body,
            string clientName, string orderRel, System.Collections.Generic.List<string> subfolders,
            out string chosenRel, out string reason)
        {
            chosenRel = null; reason = null;
            try
            {
                string b = body ?? ""; if (b.Length > 24000) b = b.Substring(0, 24000);
                string prompt =
                    "An email is being filed inside a customer ORDER folder. This may be a THREAD with several " +
                    "back-and-forth messages (newest at the top) — read all of it, then decide two things about " +
                    "the LATEST message (the one being filed), using the thread for context.\n" +
                    "1) type — is it about an ORDER (order confirmation/details) or a QUOTATION (a quote, offer " +
                    "or pricing)?\n" +
                    "2) party — who is the correspondence with, judged by the ORIGINAL correspondent's email " +
                    "DOMAIN (for a forwarded/threaded email, the external person in the conversation, NOT the " +
                    "internal colleague who forwarded it):\n" +
                    "   MI = INTERNAL — the correspondent's domain is one of OUR OWN group domains: " + InternalDomains + "\n" +
                    "   MC = the CUSTOMER — the correspondent is this order's client (" + clientName + ")\n" +
                    "   MS = a SUPPLIER — any other external company (domain not ours and not the customer)\n" +
                    "Reply with ONLY JSON: {\"type\":\"Order|Quotation\",\"party\":\"MC|MI|MS\"}\n\n" +
                    "Visible sender: " + sender + " <" + (senderEmail ?? "") + ">\nSubject: " + subject + "\n\n" + b;
                string text = ModelComplete(prompt, 0);
                var m = System.Text.RegularExpressions.Regex.Match(text ?? "", "\\{[\\s\\S]*\\}");
                if (!m.Success) return;
                var js = new System.Web.Script.Serialization.JavaScriptSerializer();
                var d = js.DeserializeObject(m.Value) as System.Collections.Generic.Dictionary<string, object>;
                if (d == null) return;
                string type = d.ContainsKey("type") && d["type"] != null ? d["type"].ToString().Trim() : "";
                string party = d.ContainsKey("party") && d["party"] != null ? d["party"].ToString().Trim() : "";
                if (party.Length == 0) return;

                // Find the real subfolder for this party. Prefer one under the right type grouping
                // (e.g. Order\MC), else any folder whose leaf is the party (e.g. MC directly under the order).
                string best = null;
                foreach (var rel in subfolders)
                {
                    if (!rel.StartsWith(orderRel + "\\", StringComparison.OrdinalIgnoreCase)) continue;
                    var segs = rel.Split('\\');
                    if (!string.Equals(segs[segs.Length - 1], party, StringComparison.OrdinalIgnoreCase)) continue;
                    bool underType = type.Length > 0 && rel.IndexOf("\\" + type + "\\", StringComparison.OrdinalIgnoreCase) >= 0;
                    if (underType) { best = rel; break; }
                    if (best == null) best = rel;
                }
                if (best == null) return;
                chosenRel = best;
                reason = party.Equals("MC", StringComparison.OrdinalIgnoreCase) ? "customer"
                       : party.Equals("MI", StringComparison.OrdinalIgnoreCase) ? "internal"
                       : party.Equals("MS", StringComparison.OrdinalIgnoreCase) ? "supplier" : party;
            }
            catch { }
        }

        // Return the immediate child of `dir` best matching `needle`: exact name first, then starts-with,
        // then contains (all case-insensitive). One directory listing; null if none / dir unreadable.
        private static string FindChild(string dir, string needle, bool exactOnly)
        {
            if (string.IsNullOrWhiteSpace(dir) || string.IsNullOrWhiteSpace(needle)) return null;
            string[] subs; try { subs = Directory.GetDirectories(dir); } catch { return null; }
            string starts = null, contains = null;
            foreach (var s in subs)
            {
                string n = Path.GetFileName(s);
                if (string.Equals(n, needle, StringComparison.OrdinalIgnoreCase)) return s;
                if (starts == null && n.StartsWith(needle, StringComparison.OrdinalIgnoreCase)) starts = s;
                if (contains == null && n.IndexOf(needle, StringComparison.OrdinalIgnoreCase) >= 0) contains = s;
            }
            return exactOnly ? null : (starts ?? contains);
        }

        // Bounded search for the first folder (within maxDepth of dir) whose name contains needle. Used
        // only as a fallback when an intermediate level (usually the client) wasn't identified, so we can
        // still locate the order by its number. Stops on the first match, at maxDepth, or when the time
        // budget is spent — never a broad, deep scan.
        private static string FindDescendant(string dir, string needle, int maxDepth,
            System.Diagnostics.Stopwatch sw, int budgetMs)
        {
            if (string.IsNullOrWhiteSpace(dir) || string.IsNullOrWhiteSpace(needle) || maxDepth <= 0
                || sw.ElapsedMilliseconds > budgetMs) return null;
            string[] subs; try { subs = Directory.GetDirectories(dir); } catch { return null; }
            foreach (var s in subs)
                if (Path.GetFileName(s).IndexOf(needle, StringComparison.OrdinalIgnoreCase) >= 0) return s;
            foreach (var s in subs)
            {
                var hit = FindDescendant(s, needle, maxDepth - 1, sw, budgetMs);
                if (hit != null) return hit;
                if (sw.ElapsedMilliseconds > budgetMs) return null;
            }
            return null;
        }

        private static string RelOf(string baseDir, string abs)
        {
            try
            {
                string b = baseDir.TrimEnd('\\', '/');
                if (abs != null && abs.Length > b.Length && abs.StartsWith(b, StringComparison.OrdinalIgnoreCase))
                    return abs.Substring(b.Length).TrimStart('\\', '/');
            }
            catch { }
            return "";
        }

        // Add every descendant of `dir` (within maxDepth) to outList as a path relative to baseDir.
        private static void CollectSubtreeRel(string baseDir, string dir, int maxDepth,
            System.Collections.Generic.List<string> outList, System.Diagnostics.Stopwatch sw, int budgetMs)
        {
            CollectSubtree(baseDir, dir, 0, maxDepth, outList, sw, budgetMs);
        }
        private static void CollectSubtree(string baseDir, string dir, int depth, int maxDepth,
            System.Collections.Generic.List<string> outList, System.Diagnostics.Stopwatch sw, int budgetMs)
        {
            if (depth >= maxDepth || outList.Count > 400 || sw.ElapsedMilliseconds > budgetMs) return;
            string[] subs; try { subs = Directory.GetDirectories(dir); } catch { return; }
            foreach (var s in subs)
            {
                AddRel(baseDir, s, outList);
                CollectSubtree(baseDir, s, depth + 1, maxDepth, outList, sw, budgetMs);
                if (outList.Count > 400 || sw.ElapsedMilliseconds > budgetMs) return;
            }
        }

        private static void AddRel(string baseDir, string abs, System.Collections.Generic.List<string> outList)
        {
            try
            {
                string b = baseDir.TrimEnd('\\', '/');
                if (abs.Length > b.Length && abs.StartsWith(b, StringComparison.OrdinalIgnoreCase))
                {
                    string rel = abs.Substring(b.Length).TrimStart('\\', '/');
                    if (rel.Length > 0 && !outList.Contains(rel)) outList.Add(rel);
                }
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
                    "Decide where THIS email should be saved.\n" +
                    "STRONGEST SIGNAL: if the subject contains an order or SAP number, the correct folder is almost " +
                    "always the one whose NAME contains that SAME number, even when the folder name has extra text " +
                    "after it (a folder is often named like '<number>_<description>'). Match the number, not the words " +
                    "around it, and include that folder's own subfolders as candidates.\n" +
                    "Otherwise use the client/company and what the email is ABOUT (the project or product named in the " +
                    "subject). List the best-fitting folders first (up to 5), copied EXACTLY from the list, and ONLY " +
                    "genuinely good fits (do not pad). " +
                    "Prefer the DEEPEST, most specific folder (the order's own leaf subfolder) over a generic parent " +
                    "folder. If nothing fits well, propose a short NEW folder path that follows the SAME structure as " +
                    "the existing folders.\n" +
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
            if (depth >= maxDepth || outList.Count > 2000) return;
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
            // Act on the email in the window the user is ACTUALLY looking at. ActiveWindow() is the
            // frontmost window — an Inspector (open email) OR the Explorer (reading pane). We deliberately
            // do NOT lead with ActiveInspector(), because that returns any email left open in a background
            // window from earlier, so Move/Reply/Summarize would silently act on the wrong email (e.g.
            // file the fans email you're reading using a stale "Sick leave" window that's still open).
            try
            {
                dynamic win = app.ActiveWindow();
                if (win != null)
                {
                    // An Inspector exposes CurrentItem; an Explorer exposes Selection. Try both, in that order.
                    try { dynamic item = win.CurrentItem; if (item != null && (int)item.Class == 43) return item; } catch { }
                    try { dynamic sel = win.Selection; if (sel != null && (int)sel.Count >= 1) { dynamic it = sel.Item(1); if ((int)it.Class == 43) return it; } } catch { }
                }
            }
            catch { }
            // Fallbacks if ActiveWindow is unavailable.
            try
            {
                dynamic exp = app.ActiveExplorer();
                if (exp != null) { dynamic sel = exp.Selection; if (sel != null && (int)sel.Count >= 1) { dynamic item = sel.Item(1); if ((int)item.Class == 43) return item; } }
            }
            catch { }
            try
            {
                dynamic insp = app.ActiveInspector();
                if (insp != null) { dynamic item = insp.CurrentItem; if (item != null && (int)item.Class == 43) return item; }
            }
            catch { }
            return null;
        }
    }
}
