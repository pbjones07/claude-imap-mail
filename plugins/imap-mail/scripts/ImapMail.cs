// ImapMail - IMAP/SMTP/MIME helpers for the ImapMail PowerShell module.
// MIT License. See LICENSE.
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Net;
using System.Net.Security;
using System.Net.Sockets;
using System.Security.Authentication;
using System.Text;
using System.Text.RegularExpressions;

namespace ImapMail
{
    public class ImapResponse
    {
        public string Text;
        public List<byte[]> Literals = new List<byte[]>();
    }

    public class FolderInfo
    {
        public string Name;
        public string Attributes;
        public int Messages = -1;
        public int Unread = -1;
    }

    public class MailSummary
    {
        public long Uid;
        public string Folder;
        public DateTime Date;
        public string From;
        public string To;
        public string Subject;
        public bool Unread;
        public bool Flagged;
        public bool Answered;
        public int SizeKB;
        public string MessageId;
    }

    public class MailAttachment
    {
        public string Name;
        public string ContentType;
        public int Size;
        public byte[] Data;
    }

    public class ParsedMail
    {
        public Dictionary<string, string> Headers = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        public string TextBody;
        public string HtmlBody;
        public List<MailAttachment> Attachments = new List<MailAttachment>();
        public string Header(string n) { string v; return Headers.TryGetValue(n, out v) ? v : null; }
        public string BestText()
        {
            if (!string.IsNullOrEmpty(TextBody) && TextBody.Trim().Length > 0) return TextBody.Replace("\r\n", "\n").Trim();
            if (HtmlBody != null) return Mime.HtmlToText(HtmlBody);
            return TextBody ?? "";
        }
    }

    public class LineConn : IDisposable
    {
        protected TcpClient tcp;
        protected Stream stream;
        byte[] buf = new byte[65536];
        int len = 0, pos = 0;

        protected void Open(string host, int port, bool tls)
        {
            tcp = new TcpClient();
            tcp.ReceiveTimeout = 90000;
            tcp.SendTimeout = 90000;
            tcp.Connect(host, port);
            stream = tcp.GetStream();
            if (tls) StartTls(host);
        }

        protected void StartTls(string host)
        {
            var s = new SslStream(stream, false);
            s.AuthenticateAsClient(host, null, SslProtocols.Tls12, true);
            stream = s; len = 0; pos = 0;
        }

        void Fill()
        {
            len = stream.Read(buf, 0, buf.Length); pos = 0;
            if (len <= 0) throw new IOException("Connection closed by server");
        }

        protected string ReadLine()
        {
            var ms = new MemoryStream();
            while (true)
            {
                if (pos >= len) Fill();
                byte b = buf[pos++];
                if (b == 10) break;
                ms.WriteByte(b);
            }
            return Encoding.UTF8.GetString(ms.ToArray()).TrimEnd('\r');
        }

        protected byte[] ReadN(int n)
        {
            var r = new byte[n]; int i = 0;
            while (i < n)
            {
                if (pos >= len) Fill();
                int c = Math.Min(n - i, len - pos);
                Buffer.BlockCopy(buf, pos, r, i, c);
                pos += c; i += c;
            }
            return r;
        }

        protected void Write(byte[] d) { stream.Write(d, 0, d.Length); stream.Flush(); }
        protected void Write(string s) { Write(Encoding.UTF8.GetBytes(s)); }

        public virtual void Dispose()
        {
            try { if (stream != null) stream.Dispose(); } catch { }
            try { if (tcp != null) tcp.Close(); } catch { }
        }
    }

    public class ImapClient : LineConn
    {
        int tag = 0;
        public string Selected;
        static readonly Regex LitRx = new Regex(@"\{(\d+)\+?\}$");
        static readonly Regex UidRx = new Regex(@"\bUID (\d+)");
        static readonly Regex FlagsRx = new Regex(@"\bFLAGS \(([^)]*)\)");
        static readonly Regex SizeRx = new Regex(@"RFC822\.SIZE (\d+)");
        static readonly Regex IdateRx = new Regex("INTERNALDATE \"([^\"]+)\"");
        static readonly Regex ListRx = new Regex("^\\* LIST \\(([^)]*)\\) (?:\"(?:[^\"\\\\]|\\\\.)*\"|NIL) (.+)$");

        public ImapClient(string host, int port)
        {
            Open(host, port, true);
            string g = ReadLine();
            if (!g.StartsWith("* OK")) throw new Exception("Unexpected greeting: " + g);
        }

        public static string Q(string s) { return "\"" + s.Replace("\\", "\\\\").Replace("\"", "\\\"") + "\""; }

        public void Login(string user, string pass) { Command("LOGIN " + Q(user) + " " + Q(pass)); }

        public List<ImapResponse> Command(string cmd) { return Command(cmd, null); }

        public List<ImapResponse> Command(string cmd, byte[] literal)
        {
            tag++;
            string t = "A" + tag.ToString("D4");
            var list = new List<ImapResponse>();
            if (literal != null)
            {
                Write(t + " " + cmd + " {" + literal.Length + "}\r\n");
                while (true)
                {
                    string l = ReadLine();
                    if (l.StartsWith("+")) break;
                    if (l.StartsWith(t + " ")) throw new Exception("IMAP error: " + l.Substring(t.Length + 1));
                }
                Write(literal);
                Write("\r\n");
            }
            else Write(t + " " + cmd + "\r\n");

            while (true)
            {
                var r = new ImapResponse();
                var sb = new StringBuilder();
                string line = ReadLine();
                while (true)
                {
                    var m = LitRx.Match(line);
                    if (m.Success)
                    {
                        sb.Append(line.Substring(0, m.Index)).Append("{L").Append(r.Literals.Count).Append("}");
                        r.Literals.Add(ReadN(int.Parse(m.Groups[1].Value)));
                        line = ReadLine();
                    }
                    else { sb.Append(line); break; }
                }
                r.Text = sb.ToString();
                list.Add(r);
                if (r.Text.StartsWith(t + " "))
                {
                    if (!r.Text.StartsWith(t + " OK")) throw new Exception("IMAP error: " + r.Text.Substring(t.Length + 1));
                    return list;
                }
            }
        }

        public void Select(string folder, bool readOnly)
        {
            Command((readOnly ? "EXAMINE " : "SELECT ") + Q(folder));
            Selected = folder;
        }

        public List<FolderInfo> ListFolders()
        {
            var res = Command("LIST \"\" \"*\"");
            var list = new List<FolderInfo>();
            foreach (var r in res)
            {
                var m = ListRx.Match(r.Text);
                if (!m.Success) continue;
                string name = m.Groups[2].Value.Trim();
                if (name.StartsWith("{L") && r.Literals.Count > 0) name = Encoding.UTF8.GetString(r.Literals[0]);
                else if (name.StartsWith("\"")) name = name.Substring(1, name.Length - 2).Replace("\\\"", "\"").Replace("\\\\", "\\");
                list.Add(new FolderInfo { Name = name, Attributes = m.Groups[1].Value });
            }
            return list;
        }

        public void FillStatus(FolderInfo f)
        {
            try
            {
                var res = Command("STATUS " + Q(f.Name) + " (MESSAGES UNSEEN)");
                foreach (var r in res)
                {
                    var m1 = Regex.Match(r.Text, @"MESSAGES (\d+)"); if (m1.Success) f.Messages = int.Parse(m1.Groups[1].Value);
                    var m2 = Regex.Match(r.Text, @"UNSEEN (\d+)"); if (m2.Success) f.Unread = int.Parse(m2.Groups[1].Value);
                }
            }
            catch { }
        }

        public string FindSpecialFolder(string attr, string fallback)
        {
            foreach (var f in ListFolders())
                if (f.Attributes.IndexOf(attr, StringComparison.OrdinalIgnoreCase) >= 0) return f.Name;
            return fallback;
        }

        public List<long> SearchUids(string criteria)
        {
            var res = Command("UID SEARCH " + criteria);
            var uids = new List<long>();
            foreach (var r in res)
                if (r.Text.StartsWith("* SEARCH"))
                    foreach (var p in r.Text.Substring(8).Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries))
                    { long u; if (long.TryParse(p, out u)) uids.Add(u); }
            uids.Sort();
            return uids;
        }

        public static string UidSet(IEnumerable<long> uids)
        {
            var l = new List<string>();
            foreach (var u in uids) l.Add(u.ToString());
            return string.Join(",", l);
        }

        public List<MailSummary> FetchSummaries(IList<long> uids)
        {
            var list = new List<MailSummary>();
            for (int i = 0; i < uids.Count; i += 200)
            {
                var chunk = new List<long>();
                for (int j = i; j < Math.Min(i + 200, uids.Count); j++) chunk.Add(uids[j]);
                var res = Command("UID FETCH " + UidSet(chunk) + " (UID FLAGS RFC822.SIZE INTERNALDATE BODY.PEEK[HEADER.FIELDS (FROM TO SUBJECT MESSAGE-ID)])");
                foreach (var r in res)
                {
                    if (!r.Text.StartsWith("* ") || r.Text.IndexOf(" FETCH ", StringComparison.Ordinal) < 0) continue;
                    var m = UidRx.Match(r.Text);
                    if (!m.Success || r.Literals.Count == 0) continue;
                    var s = new MailSummary { Folder = Selected, Uid = long.Parse(m.Groups[1].Value) };
                    m = FlagsRx.Match(r.Text);
                    string fl = m.Success ? m.Groups[1].Value : "";
                    s.Unread = fl.IndexOf("\\Seen", StringComparison.OrdinalIgnoreCase) < 0;
                    s.Flagged = fl.IndexOf("\\Flagged", StringComparison.OrdinalIgnoreCase) >= 0;
                    s.Answered = fl.IndexOf("\\Answered", StringComparison.OrdinalIgnoreCase) >= 0;
                    m = SizeRx.Match(r.Text); if (m.Success) s.SizeKB = (int)Math.Ceiling(long.Parse(m.Groups[1].Value) / 1024.0);
                    m = IdateRx.Match(r.Text); if (m.Success) s.Date = Mime.ParseImapDate(m.Groups[1].Value);
                    var h = Mime.ParseHeaders(r.Literals[0]);
                    string v;
                    if (h.TryGetValue("From", out v)) s.From = v;
                    if (h.TryGetValue("To", out v)) s.To = v;
                    if (h.TryGetValue("Subject", out v)) s.Subject = v;
                    if (h.TryGetValue("Message-ID", out v)) s.MessageId = v;
                    list.Add(s);
                }
            }
            list.Sort((a, b) => b.Date.CompareTo(a.Date));
            return list;
        }

        public byte[] FetchRaw(long uid)
        {
            var res = Command("UID FETCH " + uid + " (UID BODY.PEEK[])");
            foreach (var r in res)
                if (r.Text.IndexOf(" FETCH ", StringComparison.Ordinal) >= 0 && r.Literals.Count > 0 && UidRx.IsMatch(r.Text))
                    return r.Literals[r.Literals.Count - 1];
            return null;
        }

        public long Append(string folder, string flags, byte[] raw)
        {
            var res = Command("APPEND " + Q(folder) + " (" + flags + ")", raw);
            var m = Regex.Match(res[res.Count - 1].Text, @"APPENDUID \d+ (\d+)");
            return m.Success ? long.Parse(m.Groups[1].Value) : -1;
        }

        public void Store(IEnumerable<long> uids, string op) { Command("UID STORE " + UidSet(uids) + " " + op); }

        public void Move(IEnumerable<long> uids, string dest)
        {
            string set = UidSet(uids);
            try { Command("UID MOVE " + set + " " + Q(dest)); }
            catch
            {
                Command("UID COPY " + set + " " + Q(dest));
                Command("UID STORE " + set + " +FLAGS.SILENT (\\Deleted)");
                Command("UID EXPUNGE " + set);
            }
        }

        public void Expunge(IEnumerable<long> uids)
        {
            string set = UidSet(uids);
            Command("UID STORE " + set + " +FLAGS.SILENT (\\Deleted)");
            Command("UID EXPUNGE " + set);
        }

        public override void Dispose()
        {
            try { Command("LOGOUT"); } catch { }
            base.Dispose();
        }
    }

    public class SmtpSender : LineConn
    {
        string Resp(out int code)
        {
            string line; var sb = new StringBuilder();
            do { line = ReadLine(); sb.AppendLine(line); } while (line.Length > 3 && line[3] == '-');
            code = line.Length >= 3 ? int.Parse(line.Substring(0, 3)) : 0;
            return sb.ToString();
        }

        void Expect(string cmd, int want, string label)
        {
            if (cmd != null) Write(cmd + "\r\n");
            int c; string r = Resp(out c);
            if (c / 100 != want / 100) throw new Exception("SMTP error at " + label + ": " + r.Trim());
        }

        static byte[] DotStuff(byte[] raw)
        {
            var l1 = Encoding.GetEncoding(28591);
            string s = l1.GetString(raw);
            s = Regex.Replace(s, "\r?\n", "\r\n");
            if (!s.EndsWith("\r\n")) s += "\r\n";
            s = Regex.Replace(s, "(?m)^\\.", "..");
            return l1.GetBytes(s);
        }

        public static void Send(string host, int port, bool implicitTls, string user, string pass, string from, IEnumerable<string> rcpts, byte[] raw)
        {
            var s = new SmtpSender();
            try
            {
                s.Open(host, port, implicitTls);
                s.Expect(null, 220, "connect");
                s.Expect("EHLO localhost", 250, "EHLO");
                if (!implicitTls)
                {
                    s.Expect("STARTTLS", 220, "STARTTLS");
                    s.StartTls(host);
                    s.Expect("EHLO localhost", 250, "EHLO");
                }
                string auth = Convert.ToBase64String(Encoding.UTF8.GetBytes("\0" + user + "\0" + pass));
                s.Write("AUTH PLAIN " + auth + "\r\n");
                int c; string r = s.Resp(out c);
                if (c != 235) throw new Exception("SMTP authentication failed: " + r.Trim());
                s.Expect("MAIL FROM:<" + from + ">", 250, "MAIL FROM");
                int n = 0;
                foreach (var a in rcpts) { s.Expect("RCPT TO:<" + a + ">", 250, "RCPT TO " + a); n++; }
                if (n == 0) throw new Exception("No recipients");
                s.Expect("DATA", 354, "DATA");
                s.Write(DotStuff(raw));
                s.Expect(".", 250, "end of message");
                try { s.Write("QUIT\r\n"); } catch { }
            }
            finally { s.Dispose(); }
        }
    }

    public static class Mime
    {
        static readonly Encoding Latin1 = Encoding.GetEncoding(28591);

        static Encoding Enc(string cs)
        {
            if (string.IsNullOrEmpty(cs)) return Encoding.UTF8;
            try { return Encoding.GetEncoding(cs.Trim().Trim('"')); } catch { return Encoding.UTF8; }
        }

        // Returns length of the header block including the newline that ends the last header line.
        static int HeaderEnd(byte[] raw, out int bodyStart)
        {
            if (raw.Length >= 2 && raw[0] == 13 && raw[1] == 10) { bodyStart = 2; return 0; }
            if (raw.Length >= 1 && raw[0] == 10) { bodyStart = 1; return 0; }
            for (int i = 0; i < raw.Length - 1; i++)
            {
                if (raw[i] != 10) continue;
                if (raw[i + 1] == 10) { bodyStart = i + 2; return i + 1; }
                if (raw[i + 1] == 13 && i + 2 < raw.Length && raw[i + 2] == 10) { bodyStart = i + 3; return i + 1; }
            }
            bodyStart = raw.Length; return raw.Length;
        }

        static string HeaderText(byte[] raw, int count)
        {
            try { return new UTF8Encoding(false, true).GetString(raw, 0, count); }
            catch { return Latin1.GetString(raw, 0, count); }
        }

        static Dictionary<string, string> ParseHeaderText(string h)
        {
            var d = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            h = Regex.Replace(h, "\r?\n[ \t]+", " ");
            foreach (var line in Regex.Split(h, "\r?\n"))
            {
                int c = line.IndexOf(':');
                if (c <= 0) continue;
                string n = line.Substring(0, c).Trim();
                string v = line.Substring(c + 1).Trim();
                if (!d.ContainsKey(n)) d[n] = v;
            }
            return d;
        }

        public static Dictionary<string, string> ParseHeaders(byte[] raw)
        {
            int bs; int he = HeaderEnd(raw, out bs);
            var d = ParseHeaderText(HeaderText(raw, he));
            var o = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            foreach (var kv in d) o[kv.Key] = DecodeWords(kv.Value);
            return o;
        }

        static readonly Regex EwRx = new Regex(@"=\?([^?\s]+)\?([BbQq])\?([^?\s]*)\?=");

        public static string DecodeWords(string s)
        {
            if (s == null || s.IndexOf("=?", StringComparison.Ordinal) < 0) return s;
            s = Regex.Replace(s, @"(\?=)\s+(=\?)", "$1$2");
            return EwRx.Replace(s, m =>
            {
                try
                {
                    string cs = m.Groups[1].Value;
                    int star = cs.IndexOf('*'); if (star >= 0) cs = cs.Substring(0, star);
                    byte[] b;
                    if (m.Groups[2].Value.ToUpperInvariant() == "B") b = Convert.FromBase64String(PadB64(m.Groups[3].Value));
                    else b = QDecode(m.Groups[3].Value.Replace('_', ' '), false);
                    return Enc(cs).GetString(b);
                }
                catch { return m.Value; }
            });
        }

        static string PadB64(string s)
        {
            s = Regex.Replace(s, @"[^A-Za-z0-9+/=]", "");
            s = s.TrimEnd('=');
            while (s.Length % 4 != 0) s += "=";
            return s;
        }

        static bool IsHex(char c) { return (c >= '0' && c <= '9') || (c >= 'A' && c <= 'F') || (c >= 'a' && c <= 'f'); }

        static byte[] QDecode(string s, bool qp)
        {
            if (qp) s = Regex.Replace(s, "=[ \t]*\r?\n", "");
            var ms = new MemoryStream();
            for (int i = 0; i < s.Length; i++)
            {
                char ch = s[i];
                if (ch == '=' && i + 2 < s.Length && IsHex(s[i + 1]) && IsHex(s[i + 2]))
                {
                    ms.WriteByte(Convert.ToByte(s.Substring(i + 1, 2), 16)); i += 2;
                }
                else if (ch < 256) ms.WriteByte((byte)ch);
                else { var b = Encoding.UTF8.GetBytes(ch.ToString()); ms.Write(b, 0, b.Length); }
            }
            return ms.ToArray();
        }

        static string MainValue(string v)
        {
            if (v == null) return "";
            int i = v.IndexOf(';');
            return (i < 0 ? v : v.Substring(0, i)).Trim().ToLowerInvariant();
        }

        static readonly Regex ParamRx = new Regex(";\\s*([A-Za-z0-9\\-_\\.\\*]+)\\s*=\\s*(\"((?:[^\"\\\\]|\\\\.)*)\"|[^;\\s]+)");

        static Dictionary<string, string> Params(string v)
        {
            var d = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            if (v == null) return d;
            foreach (Match m in ParamRx.Matches(v))
            {
                string k = m.Groups[1].Value;
                string val = m.Groups[3].Success ? m.Groups[3].Value.Replace("\\\"", "\"") : m.Groups[2].Value;
                if (k.EndsWith("*"))
                {
                    k = k.TrimEnd('*');
                    int q = val.IndexOf("''", StringComparison.Ordinal);
                    string cs = q > 0 ? val.Substring(0, q) : "utf-8";
                    if (q >= 0) val = val.Substring(q + 2);
                    try { val = Enc(cs).GetString(QDecode(val.Replace('%', '='), false)); } catch { }
                }
                var cm = Regex.Match(k, @"^(.*)\*\d+$");
                if (cm.Success)
                {
                    k = cm.Groups[1].Value;
                    string prev; if (d.TryGetValue(k, out prev)) val = prev + val;
                }
                d[k] = val;
            }
            return d;
        }

        static byte[] Sub(byte[] a, int start)
        {
            if (start >= a.Length) return new byte[0];
            var r = new byte[a.Length - start];
            Buffer.BlockCopy(a, start, r, 0, r.Length);
            return r;
        }

        public static ParsedMail Parse(byte[] raw)
        {
            var pm = new ParsedMail();
            int bs; int he = HeaderEnd(raw, out bs);
            var hd = ParseHeaderText(HeaderText(raw, he));
            foreach (var kv in hd) pm.Headers[kv.Key] = DecodeWords(kv.Value);
            Walk(hd, Sub(raw, bs), pm, 0);
            return pm;
        }

        static void Walk(Dictionary<string, string> hd, byte[] body, ParsedMail pm, int depth)
        {
            string ctv; hd.TryGetValue("Content-Type", out ctv);
            string ct = MainValue(ctv); if (ct == "") ct = "text/plain";
            var cp = Params(ctv);
            string cdv; hd.TryGetValue("Content-Disposition", out cdv);
            var dp = Params(cdv);
            string cte; hd.TryGetValue("Content-Transfer-Encoding", out cte);
            cte = (cte ?? "").Trim().ToLowerInvariant();

            if (ct.StartsWith("multipart/") && cp.ContainsKey("boundary") && depth < 20)
            {
                foreach (var part in SplitMultipart(body, cp["boundary"]))
                {
                    int bs; int he = HeaderEnd(part, out bs);
                    var ph = ParseHeaderText(HeaderText(part, he));
                    Walk(ph, Sub(part, bs), pm, depth + 1);
                }
                return;
            }

            byte[] data = Decode(body, cte);
            string fname = null;
            if (dp.ContainsKey("filename")) fname = dp["filename"];
            else if (cp.ContainsKey("name")) fname = cp["name"];
            bool attach = MainValue(cdv) == "attachment";

            if (!attach && (ct == "text/plain" || ct == "text/html"))
            {
                string text = Enc(cp.ContainsKey("charset") ? cp["charset"] : null).GetString(data);
                if (ct == "text/plain") { if (pm.TextBody == null) pm.TextBody = text; }
                else { if (pm.HtmlBody == null) pm.HtmlBody = text; }
                return;
            }
            pm.Attachments.Add(new MailAttachment
            {
                Name = fname != null ? DecodeWords(fname) : "(unnamed " + ct + ")",
                ContentType = ct,
                Size = data.Length,
                Data = data
            });
        }

        static byte[] Decode(byte[] body, string cte)
        {
            if (cte == "base64")
            {
                try { return Convert.FromBase64String(PadB64(Latin1.GetString(body))); } catch { return body; }
            }
            if (cte == "quoted-printable") return QDecode(Latin1.GetString(body), true);
            return body;
        }

        static List<byte[]> SplitMultipart(byte[] body, string boundary)
        {
            var parts = new List<byte[]>();
            string s = Latin1.GetString(body);
            string delim = "--" + boundary;
            StringBuilder cur = null;
            foreach (var line in Regex.Split(s, "(?<=\n)"))
            {
                string t = line.TrimEnd('\r', '\n', ' ', '\t');
                if (t == delim || t == delim + "--")
                {
                    if (cur != null) parts.Add(Latin1.GetBytes(TrimOneNewline(cur.ToString())));
                    if (t == delim + "--") { cur = null; break; }
                    cur = new StringBuilder();
                    continue;
                }
                if (cur != null) cur.Append(line);
            }
            if (cur != null) parts.Add(Latin1.GetBytes(cur.ToString()));
            return parts;
        }

        static string TrimOneNewline(string s)
        {
            if (s.EndsWith("\r\n")) return s.Substring(0, s.Length - 2);
            if (s.EndsWith("\n")) return s.Substring(0, s.Length - 1);
            return s;
        }

        public static string HtmlToText(string html)
        {
            string s = Regex.Replace(html, @"(?is)<(script|style|head|title)[^>]*>.*?</\1\s*>", "");
            s = Regex.Replace(s, @"(?s)<!--.*?-->", "");
            s = Regex.Replace(s, @"[\r\n]+", " ");
            s = Regex.Replace(s, @"(?i)<br\s*/?>", "\n");
            s = Regex.Replace(s, @"(?i)</(p|div|tr|li|h[1-6]|table|blockquote|ul|ol)\s*>", "\n");
            s = Regex.Replace(s, @"(?i)<li[^>]*>", "\u2022 ");
            s = Regex.Replace(s, @"(?i)</t[dh]\s*>", " ");
            s = Regex.Replace(s, @"<[^>]+>", "");
            s = WebUtility.HtmlDecode(s);
            s = s.Replace('\u00a0', ' ');
            s = Regex.Replace(s, "[\u200b\u200c\u200d\u034f\u00ad\ufeff\u2007]", "");
            s = Regex.Replace(s, @"[ \t]+", " ");
            s = Regex.Replace(s, @" *\n *", "\n");
            s = Regex.Replace(s, @"\n{3,}", "\n\n");
            return s.Trim();
        }

        public static DateTime ParseImapDate(string s)
        {
            s = Regex.Replace(s.Trim(), @"([+-]\d\d)(\d\d)$", "$1:$2");
            DateTimeOffset d;
            if (DateTimeOffset.TryParseExact(s, "d-MMM-yyyy HH:mm:ss zzz", CultureInfo.InvariantCulture, DateTimeStyles.None, out d)) return d.LocalDateTime;
            return DateTime.MinValue;
        }

        static bool IsAscii(string v)
        {
            foreach (char c in v) if (c > 126 || (c < 32 && c != '\t')) return false;
            return true;
        }

        public static string EncodeHeader(string v)
        {
            if (v == null) return "";
            if (IsAscii(v)) return v;
            var words = new List<string>();
            int i = 0;
            while (i < v.Length)
            {
                int n = Math.Min(20, v.Length - i);
                if (n < v.Length - i && char.IsHighSurrogate(v[i + n - 1])) n--;
                words.Add("=?utf-8?B?" + Convert.ToBase64String(Encoding.UTF8.GetBytes(v.Substring(i, n))) + "?=");
                i += n;
            }
            return string.Join("\r\n ", words);
        }

        public static string EncodeAddressList(string list)
        {
            if (string.IsNullOrWhiteSpace(list)) return "";
            var items = new List<string>();
            var sb = new StringBuilder(); bool inQ = false;
            foreach (char c in list)
            {
                if (c == '"') inQ = !inQ;
                if ((c == ',' || c == ';') && !inQ) { items.Add(sb.ToString()); sb.Clear(); continue; }
                sb.Append(c);
            }
            items.Add(sb.ToString());
            var outl = new List<string>();
            foreach (var raw in items)
            {
                string it = raw.Trim();
                if (it == "") continue;
                var m = Regex.Match(it, "^\"?([^\"<]*?)\"?\\s*<([^>]+)>$");
                if (m.Success)
                {
                    string name = m.Groups[1].Value.Trim();
                    string addr = m.Groups[2].Value.Trim();
                    if (name == "") outl.Add("<" + addr + ">");
                    else if (IsAscii(name)) outl.Add("\"" + name.Replace("\"", "") + "\" <" + addr + ">");
                    else outl.Add(EncodeHeader(name) + " <" + addr + ">");
                }
                else outl.Add(it);
            }
            return string.Join(",\r\n ", outl);
        }

        static string QPEncode(string text)
        {
            var o = new StringBuilder();
            var lines = text.Replace("\r\n", "\n").Replace("\r", "\n").Split('\n');
            foreach (var line in lines)
            {
                byte[] b = Encoding.UTF8.GetBytes(line);
                int col = 0;
                for (int i = 0; i < b.Length; i++)
                {
                    byte c = b[i];
                    bool last = i == b.Length - 1;
                    string tok;
                    if ((c >= 33 && c <= 126 && c != 61) || ((c == 32 || c == 9) && !last)) tok = ((char)c).ToString();
                    else tok = "=" + c.ToString("X2");
                    if (col + tok.Length > 75) { o.Append("=\r\n"); col = 0; }
                    o.Append(tok); col += tok.Length;
                }
                o.Append("\r\n");
            }
            return o.ToString();
        }

        public static byte[] BuildMessage(string from, string to, string cc, string bcc, string subject, string body, string inReplyTo, string references, out string messageId)
        {
            string dom = "icloud.com";
            var am = Regex.Match(from ?? "", @"@([A-Za-z0-9.\-]+)");
            if (am.Success) dom = am.Groups[1].Value;
            messageId = "<" + Guid.NewGuid().ToString("N") + "@" + dom + ">";
            var now = DateTimeOffset.Now;
            var sb = new StringBuilder();
            sb.Append("Date: ").Append(now.ToString("ddd, dd MMM yyyy HH:mm:ss ", CultureInfo.InvariantCulture)).Append(now.ToString("zzz", CultureInfo.InvariantCulture).Replace(":", "")).Append("\r\n");
            sb.Append("From: ").Append(EncodeAddressList(from)).Append("\r\n");
            if (!string.IsNullOrWhiteSpace(to)) sb.Append("To: ").Append(EncodeAddressList(to)).Append("\r\n");
            if (!string.IsNullOrWhiteSpace(cc)) sb.Append("Cc: ").Append(EncodeAddressList(cc)).Append("\r\n");
            if (!string.IsNullOrWhiteSpace(bcc)) sb.Append("Bcc: ").Append(EncodeAddressList(bcc)).Append("\r\n");
            sb.Append("Subject: ").Append(EncodeHeader(subject ?? "")).Append("\r\n");
            sb.Append("Message-ID: ").Append(messageId).Append("\r\n");
            if (!string.IsNullOrWhiteSpace(inReplyTo)) sb.Append("In-Reply-To: ").Append(inReplyTo.Trim()).Append("\r\n");
            if (!string.IsNullOrWhiteSpace(references)) sb.Append("References: ").Append(references.Trim()).Append("\r\n");
            sb.Append("MIME-Version: 1.0\r\n");
            sb.Append("Content-Type: text/plain; charset=utf-8\r\n");
            sb.Append("Content-Transfer-Encoding: quoted-printable\r\n");
            sb.Append("\r\n");
            sb.Append(QPEncode(body ?? ""));
            return Encoding.ASCII.GetBytes(sb.ToString());
        }

        public static byte[] RemoveHeader(byte[] raw, string name)
        {
            int bs; int he = HeaderEnd(raw, out bs);
            string h = Latin1.GetString(raw, 0, he);
            h = Regex.Replace(h, "(?im)^" + Regex.Escape(name) + @":[^\n]*\n(?:[ \t][^\n]*\n)*", "");
            var hb = Latin1.GetBytes(h);
            var r = new byte[hb.Length + raw.Length - he];
            Buffer.BlockCopy(hb, 0, r, 0, hb.Length);
            Buffer.BlockCopy(raw, he, r, hb.Length, raw.Length - he);
            return r;
        }
    }
}
