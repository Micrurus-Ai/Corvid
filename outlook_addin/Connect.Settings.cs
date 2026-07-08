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
                while (m != null && n < 40)
                {
                    string b = ""; try { b = (string)m.Body; } catch { }
                    foreach (var mark in new[] { "\r\nFrom:", "\r\nVan:", "\r\nSent:", "-----Original", "-----Oorspronkelijk" })
                    { int i = b.IndexOf(mark); if (i > 0) { b = b.Substring(0, i); } }
                    b = (b ?? "").Trim();
                    if (b.Length >= 20)
                    {
                        if (b.Length > 1500) b = b.Substring(0, 1500);
                        sb.AppendLine("=====EMAIL====="); sb.AppendLine(b); n++;
                    }
                    try { m = items.GetNext(); } catch { m = null; }
                }
                if (n == 0) return "";
                string prompt =
                    "Below are recent emails the USER has SENT (their own outgoing messages, quoted history removed). " +
                    "Study them and write a detailed WRITING-STYLE PROFILE so another writer can produce emails that " +
                    "are indistinguishable from this person's — capture their VOICE, not just structure:\n" +
                    "- Greetings, and how they address people (first name? 'Hi'/'Hello'/'Dear'/'Beste'/'Bonjour'?)\n" +
                    "- Sign-offs and signature — the EXACT wording they repeat\n" +
                    "- Formality and warmth, and how it shifts for colleagues vs clients\n" +
                    "- Sentence length and structure (short & direct, or longer; paragraphs vs bullets)\n" +
                    "- Recurring words, phrases and expressions they actually reuse (their pet phrases, connective " +
                    "words, how they thank/ask/confirm) — ONLY ones appearing in several emails\n" +
                    "- Punctuation habits (exclamation marks, dashes, emojis) and level of directness/politeness\n" +
                    "- Languages they write in (Dutch/French/English) and any differences per language\n" +
                    "Describe PATTERNS, not one-off examples; ignore any quoted or forwarded text from other people. " +
                    "Be specific and concrete. Plain text with '- ' bullets, no Markdown, no asterisks, about 250 words.\n\n"
                    + sb.ToString();
                string tone = ModelComplete(prompt, 0.2);
                return (tone ?? "").Replace("*", "").Trim();
            }
            catch { return ""; }
        }
    }
}
