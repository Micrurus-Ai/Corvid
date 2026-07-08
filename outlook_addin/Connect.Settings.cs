// Axon Outlook add-in — Settings dialog entry point + learn-tone-from-Sent.  (partial of Connect; split out of AxonAddin.cs.)
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
    }
}
