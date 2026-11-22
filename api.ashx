<%@ WebHandler Language="C#" Class="NoteItApi" %>
<%@ Assembly Name="System.IO.Compression, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089" %>
<%@ Assembly Name="System.Web.Extensions, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35" %>
<%@ Assembly Name="System.Xml.Linq, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089" %>
<%@ Assembly Name="System.Core, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089" %>

// ============================================================================
//  Note It  -  server API for IIS (ASP.NET 4.5+)
//  Data is kept in two Excel files inside the App_Data folder:
//     App_Data\Users.xlsx   -> sheet "Users"
//     App_Data\Tasks.xlsx   -> sheets "Notes" and "Tasks"
//  No extra libraries or database are needed.
// ============================================================================

using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Text.RegularExpressions;
using System.Web;
using System.Web.Script.Serialization;
using System.Xml.Linq;

// ---------------------------------------------------------------- data types
public class Row : Dictionary<string, string>
{
    public Row() : base(StringComparer.OrdinalIgnoreCase) { }
    public string Get(string k) { string v; return TryGetValue(k, out v) && v != null ? v.Trim() : ""; }
}

public class Sheet
{
    public string Name;
    public List<string> Headers = new List<string>();
    public List<Row> Rows = new List<Row>();
    public Sheet(string name, string[] headers) { Name = name; Headers.AddRange(headers); }
    public void EnsureHeaders(string[] required)
    {
        foreach (string h in required)
            if (!Headers.Any(x => string.Equals(x, h, StringComparison.OrdinalIgnoreCase))) Headers.Add(h);
    }
}

public class ApiError : Exception
{
    public int Status;
    public ApiError(string msg, int status) : base(msg) { Status = status; }
    public ApiError(string msg) : this(msg, 400) { }
}

// ---------------------------------------------------------------- xlsx read/write (no libraries)
public static class Xlsx
{
    static readonly XNamespace M = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
    static readonly XNamespace R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
    static readonly XNamespace PR = "http://schemas.openxmlformats.org/package/2006/relationships";

    static ZipArchiveEntry Find(ZipArchive z, string name)
    {
        name = name.TrimStart('/');
        foreach (ZipArchiveEntry e in z.Entries)
            if (string.Equals(e.FullName.TrimStart('/'), name, StringComparison.OrdinalIgnoreCase)) return e;
        return null;
    }

    static XDocument Load(ZipArchiveEntry e)
    {
        using (Stream s = e.Open()) return XDocument.Load(s);
    }

    static int ColIndex(string cellRef)
    {
        int n = 0;
        foreach (char ch in cellRef)
        {
            if (char.IsLetter(ch)) n = n * 26 + (char.ToUpperInvariant(ch) - 'A' + 1);
            else break;
        }
        return n - 1;
    }

    static string ColName(int i)
    {
        string s = "";
        i++;
        while (i > 0) { int m = (i - 1) % 26; s = (char)('A' + m) + s; i = (i - 1) / 26; }
        return s;
    }

    static string AllText(XElement el)
    {
        if (el == null) return "";
        return string.Concat(el.Descendants(M + "t").Select(t => t.Value));
    }

    /// Reads every sheet. dateCols: columns whose numeric values are Excel dates.
    public static Dictionary<string, Sheet> Read(string path, string[] dateCols)
    {
        var result = new Dictionary<string, Sheet>(StringComparer.OrdinalIgnoreCase);
        using (var fs = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite | FileShare.Delete))
        using (var zip = new ZipArchive(fs, ZipArchiveMode.Read))
        {
            var shared = new List<string>();
            ZipArchiveEntry ss = Find(zip, "xl/sharedStrings.xml");
            if (ss != null)
                foreach (XElement si in Load(ss).Root.Elements(M + "si")) shared.Add(AllText(si));

            XDocument wb = Load(Find(zip, "xl/workbook.xml"));
            XDocument rels = Load(Find(zip, "xl/_rels/workbook.xml.rels"));
            foreach (XElement sh in wb.Root.Element(M + "sheets").Elements(M + "sheet"))
            {
                string name = (string)sh.Attribute("name");
                string rid = (string)sh.Attribute(R + "id");
                XElement rel = rels.Root.Elements(PR + "Relationship").FirstOrDefault(r => (string)r.Attribute("Id") == rid);
                if (rel == null) continue;
                string target = (string)rel.Attribute("Target");
                target = target.StartsWith("/") ? target.TrimStart('/') : "xl/" + target;
                ZipArchiveEntry entry = Find(zip, target);
                if (entry == null) continue;

                var sheet = new Sheet(name, new string[0]);
                List<string> headers = null;
                XElement data = Load(entry).Root.Element(M + "sheetData");
                if (data != null)
                {
                    foreach (XElement row in data.Elements(M + "row"))
                    {
                        var vals = new Dictionary<int, string>();
                        int next = 0;
                        foreach (XElement c in row.Elements(M + "c"))
                        {
                            string r = (string)c.Attribute("r");
                            int col = string.IsNullOrEmpty(r) ? next : ColIndex(r);
                            next = col + 1;
                            string t = (string)c.Attribute("t") ?? "";
                            XElement v = c.Element(M + "v");
                            string val;
                            if (t == "s") { int idx; val = (v != null && int.TryParse(v.Value, out idx) && idx < shared.Count) ? shared[idx] : ""; }
                            else if (t == "inlineStr") val = AllText(c.Element(M + "is"));
                            else if (t == "b") val = (v != null && v.Value == "1") ? "TRUE" : "FALSE";
                            else val = v != null ? v.Value : "";
                            vals[col] = val;
                        }
                        if (headers == null)
                        {
                            headers = new List<string>();
                            int max = vals.Count == 0 ? -1 : vals.Keys.Max();
                            for (int i = 0; i <= max; i++) headers.Add(vals.ContainsKey(i) ? vals[i].Trim() : "");
                            continue;
                        }
                        var rr = new Row();
                        bool any = false;
                        for (int i = 0; i < headers.Count; i++)
                        {
                            if (headers[i] == "") continue;
                            string val = vals.ContainsKey(i) ? vals[i] : "";
                            double d;
                            if (val != "" && double.TryParse(val, NumberStyles.Float, CultureInfo.InvariantCulture, out d))
                            {
                                bool isDate = dateCols.Any(x => string.Equals(x, headers[i], StringComparison.OrdinalIgnoreCase));
                                if (isDate && d > 1 && d < 2958465)
                                {
                                    DateTime dt = DateTime.FromOADate(d);
                                    val = dt.TimeOfDay.TotalMinutes >= 1 ? dt.ToString("yyyy-MM-dd HH:mm") : dt.ToString("yyyy-MM-dd");
                                }
                                else val = d.ToString("0.###############", CultureInfo.InvariantCulture);
                            }
                            rr[headers[i]] = val;
                            if (val.Trim() != "") any = true;
                        }
                        if (any) sheet.Rows.Add(rr);
                    }
                }
                if (headers != null) sheet.Headers = headers.Where(h => h != "").ToList();
                result[name] = sheet;
            }
        }
        return result;
    }

    static string Clean(string s)
    {
        if (s == null) return "";
        var sb = new StringBuilder(s.Length);
        foreach (char ch in s) if (ch == '\t' || ch == '\n' || ch == '\r' || ch >= ' ') sb.Append(ch);
        return sb.ToString();
    }

    static void Put(ZipArchive z, string name, XDocument doc)
    {
        ZipArchiveEntry e = z.CreateEntry(name, CompressionLevel.Optimal);
        using (Stream s = e.Open())
        using (var w = new StreamWriter(s, new UTF8Encoding(false)))
            doc.Save(w, SaveOptions.DisableFormatting);
    }

    /// Writes all sheets to a new xlsx file (header row bold and frozen). Writes to a temp file first.
    public static void Write(string path, IList<Sheet> sheets)
    {
        string tmp = path + ".tmp";
        using (var fs = new FileStream(tmp, FileMode.Create, FileAccess.Write))
        using (var z = new ZipArchive(fs, ZipArchiveMode.Create))
        {
            XNamespace CT = "http://schemas.openxmlformats.org/package/2006/content-types";
            var types = new XElement(CT + "Types",
                new XElement(CT + "Default", new XAttribute("Extension", "rels"), new XAttribute("ContentType", "application/vnd.openxmlformats-package.relationships+xml")),
                new XElement(CT + "Default", new XAttribute("Extension", "xml"), new XAttribute("ContentType", "application/xml")),
                new XElement(CT + "Override", new XAttribute("PartName", "/xl/workbook.xml"), new XAttribute("ContentType", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml")),
                new XElement(CT + "Override", new XAttribute("PartName", "/xl/styles.xml"), new XAttribute("ContentType", "application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml")));
            for (int i = 0; i < sheets.Count; i++)
                types.Add(new XElement(CT + "Override", new XAttribute("PartName", "/xl/worksheets/sheet" + (i + 1) + ".xml"), new XAttribute("ContentType", "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml")));
            Put(z, "[Content_Types].xml", new XDocument(types));

            Put(z, "_rels/.rels", new XDocument(new XElement(PR + "Relationships",
                new XElement(PR + "Relationship", new XAttribute("Id", "rId1"),
                    new XAttribute("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"),
                    new XAttribute("Target", "xl/workbook.xml")))));

            var sheetsEl = new XElement(M + "sheets");
            var wbRels = new XElement(PR + "Relationships");
            for (int i = 0; i < sheets.Count; i++)
            {
                sheetsEl.Add(new XElement(M + "sheet", new XAttribute("name", sheets[i].Name), new XAttribute("sheetId", i + 1), new XAttribute(R + "id", "rId" + (i + 1))));
                wbRels.Add(new XElement(PR + "Relationship", new XAttribute("Id", "rId" + (i + 1)),
                    new XAttribute("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet"),
                    new XAttribute("Target", "worksheets/sheet" + (i + 1) + ".xml")));
            }
            wbRels.Add(new XElement(PR + "Relationship", new XAttribute("Id", "rId" + (sheets.Count + 1)),
                new XAttribute("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles"),
                new XAttribute("Target", "styles.xml")));
            Put(z, "xl/workbook.xml", new XDocument(new XElement(M + "workbook", new XAttribute(XNamespace.Xmlns + "r", R.NamespaceName), sheetsEl)));
            Put(z, "xl/_rels/workbook.xml.rels", new XDocument(wbRels));

            Put(z, "xl/styles.xml", new XDocument(new XElement(M + "styleSheet",
                new XElement(M + "fonts", new XAttribute("count", 2),
                    new XElement(M + "font", new XElement(M + "sz", new XAttribute("val", 11)), new XElement(M + "name", new XAttribute("val", "Calibri"))),
                    new XElement(M + "font", new XElement(M + "b"), new XElement(M + "color", new XAttribute("rgb", "FFFFFFFF")), new XElement(M + "sz", new XAttribute("val", 11)), new XElement(M + "name", new XAttribute("val", "Calibri")))),
                new XElement(M + "fills", new XAttribute("count", 3),
                    new XElement(M + "fill", new XElement(M + "patternFill", new XAttribute("patternType", "none"))),
                    new XElement(M + "fill", new XElement(M + "patternFill", new XAttribute("patternType", "gray125"))),
                    new XElement(M + "fill", new XElement(M + "patternFill", new XAttribute("patternType", "solid"),
                        new XElement(M + "fgColor", new XAttribute("rgb", "FF4F46E5")), new XElement(M + "bgColor", new XAttribute("indexed", 64))))),
                new XElement(M + "borders", new XAttribute("count", 1), new XElement(M + "border", new XElement(M + "left"), new XElement(M + "right"), new XElement(M + "top"), new XElement(M + "bottom"), new XElement(M + "diagonal"))),
                new XElement(M + "cellStyleXfs", new XAttribute("count", 1), new XElement(M + "xf", new XAttribute("numFmtId", 0), new XAttribute("fontId", 0), new XAttribute("fillId", 0), new XAttribute("borderId", 0))),
                new XElement(M + "cellXfs", new XAttribute("count", 2),
                    new XElement(M + "xf", new XAttribute("numFmtId", 0), new XAttribute("fontId", 0), new XAttribute("fillId", 0), new XAttribute("borderId", 0), new XAttribute("xfId", 0)),
                    new XElement(M + "xf", new XAttribute("numFmtId", 0), new XAttribute("fontId", 1), new XAttribute("fillId", 2), new XAttribute("borderId", 0), new XAttribute("xfId", 0), new XAttribute("applyFont", 1), new XAttribute("applyFill", 1))),
                new XElement(M + "cellStyles", new XAttribute("count", 1), new XElement(M + "cellStyle", new XAttribute("name", "Normal"), new XAttribute("xfId", 0), new XAttribute("builtinId", 0))))));

            for (int i = 0; i < sheets.Count; i++)
            {
                Sheet sh = sheets[i];
                var cols = new XElement(M + "cols");
                for (int c = 0; c < sh.Headers.Count; c++)
                {
                    int w = sh.Headers[c].Length;
                    foreach (Row r in sh.Rows) w = Math.Max(w, Math.Min(r.Get(sh.Headers[c]).Length, 60));
                    cols.Add(new XElement(M + "col", new XAttribute("min", c + 1), new XAttribute("max", c + 1), new XAttribute("width", Math.Max(10, w + 3)), new XAttribute("customWidth", 1)));
                }
                var sd = new XElement(M + "sheetData");
                var all = new List<List<string>>();
                all.Add(sh.Headers);
                foreach (Row r in sh.Rows) all.Add(sh.Headers.Select(h => r.Get(h)).ToList());
                for (int ri = 0; ri < all.Count; ri++)
                {
                    var rowEl = new XElement(M + "row", new XAttribute("r", ri + 1));
                    for (int c = 0; c < all[ri].Count; c++)
                    {
                        string v = Clean(all[ri][c]);
                        if (v == "" && ri > 0) continue;
                        var cell = new XElement(M + "c", new XAttribute("r", ColName(c) + (ri + 1)), new XAttribute("t", "inlineStr"));
                        if (ri == 0) cell.Add(new XAttribute("s", 1));
                        var tEl = new XElement(M + "t", v);
                        if (v != v.Trim()) tEl.Add(new XAttribute(XNamespace.Xml + "space", "preserve"));
                        cell.Add(new XElement(M + "is", tEl));
                        rowEl.Add(cell);
                    }
                    sd.Add(rowEl);
                }
                var views = new XElement(M + "sheetViews", new XElement(M + "sheetView", new XAttribute("workbookViewId", 0),
                    new XElement(M + "pane", new XAttribute("ySplit", 1), new XAttribute("topLeftCell", "A2"), new XAttribute("activePane", "bottomLeft"), new XAttribute("state", "frozen"))));
                Put(z, "xl/worksheets/sheet" + (i + 1) + ".xml", new XDocument(new XElement(M + "worksheet", views, cols, sd)));
            }
        }
        File.Copy(tmp, path, true);   // throws IOException if Excel has the file open
        try { File.Delete(tmp); } catch { }
    }
}

// ---------------------------------------------------------------- storage
public static class Store
{
    public static readonly object Lock = new object();
    public static readonly string[] UserCols = { "UserID", "Name", "Password", "Role", "Active", "Created" };
    public static readonly string[] NoteCols = { "NoteID", "Title", "CreatedBy", "Members", "Created", "Updated" };
    public static readonly string[] TaskCols = { "TaskID", "NoteID", "Task", "AssignedTo", "DueDate", "Done", "AddedBy", "Created" };
    static readonly string[] DateCols = { "Created", "Updated", "DueDate" };

    public static string DataDirOverride;
    public static string Dir { get { return DataDirOverride ?? HttpContext.Current.Server.MapPath("~/App_Data"); } }
    public static string UsersPath { get { return Path.Combine(Dir, "Users.xlsx"); } }
    public static string TasksPath { get { return Path.Combine(Dir, "Tasks.xlsx"); } }

    public static Sheet Users, Notes, Tasks;
    static DateTime usersStamp, tasksStamp;

    public static string Now() { return DateTime.Now.ToString("yyyy-MM-dd HH:mm"); }

    static Sheet Pick(Dictionary<string, Sheet> book, string name, string[] cols)
    {
        Sheet s;
        if (!book.TryGetValue(name, out s))
            s = book.Count > 0 && name == "Users" ? book.Values.First() : new Sheet(name, cols);
        s.Name = name;
        s.EnsureHeaders(cols);
        return s;
    }

    /// Loads both files (only re-reads a file when it has changed on disk). Call inside lock.
    public static void Load()
    {
        if (!Directory.Exists(Dir)) Directory.CreateDirectory(Dir);
        if (!File.Exists(UsersPath))
        {
            var u = new Sheet("Users", UserCols);
            var admin = new Row();
            admin["UserID"] = "admin"; admin["Name"] = "Administrator"; admin["Password"] = "admin123";
            admin["Role"] = "Admin"; admin["Active"] = "Yes"; admin["Created"] = Now();
            u.Rows.Add(admin);
            Xlsx.Write(UsersPath, new List<Sheet> { u });
        }
        if (!File.Exists(TasksPath))
            Xlsx.Write(TasksPath, new List<Sheet> { new Sheet("Notes", NoteCols), new Sheet("Tasks", TaskCols) });

        DateTime us = File.GetLastWriteTimeUtc(UsersPath);
        if (Users == null || us != usersStamp)
        {
            Users = Pick(Xlsx.Read(UsersPath, DateCols), "Users", UserCols);
            usersStamp = us;
        }
        DateTime ts = File.GetLastWriteTimeUtc(TasksPath);
        if (Notes == null || ts != tasksStamp)
        {
            var book = Xlsx.Read(TasksPath, DateCols);
            Notes = Pick(book, "Notes", NoteCols);
            Tasks = Pick(book, "Tasks", TaskCols);
            tasksStamp = ts;
        }
    }

    public static void SaveUsers()
    {
        try { Xlsx.Write(UsersPath, new List<Sheet> { Users }); }
        catch (IOException) { Users = null; throw new ApiError("Users.xlsx is open in Excel (or locked). Close it and try again.", 409); }
        usersStamp = File.GetLastWriteTimeUtc(UsersPath);
    }

    public static void SaveTasks()
    {
        try { Xlsx.Write(TasksPath, new List<Sheet> { Notes, Tasks }); }
        catch (IOException) { Notes = null; throw new ApiError("Tasks.xlsx is open in Excel (or locked). Close it and try again.", 409); }
        tasksStamp = File.GetLastWriteTimeUtc(TasksPath);
    }

    public static string NextId(Sheet s, string col, string prefix)
    {
        int max = 0;
        foreach (Row r in s.Rows)
        {
            string id = r.Get(col);
            int n;
            if (id.StartsWith(prefix, StringComparison.OrdinalIgnoreCase) && int.TryParse(id.Substring(prefix.Length), out n) && n > max) max = n;
        }
        return prefix + (max + 1).ToString("00000");
    }
}

// ---------------------------------------------------------------- security helpers
public static class Sec
{
    public static bool Same(string a, string b) { return string.Equals((a ?? "").Trim(), (b ?? "").Trim(), StringComparison.OrdinalIgnoreCase); }

    static string Hex(byte[] b) { return BitConverter.ToString(b).Replace("-", "").ToLowerInvariant(); }

    public static string Hash(string password)
    {
        byte[] salt = new byte[16];
        using (var rng = RandomNumberGenerator.Create()) rng.GetBytes(salt);
        string s = Hex(salt);
        return "sha256$" + s + "$" + Sha(s, password);
    }

    static string Sha(string salt, string password)
    {
        using (var sha = SHA256.Create())
            return Hex(sha.ComputeHash(Encoding.UTF8.GetBytes(salt + ":" + password)));
    }

    public static bool IsHashed(string stored) { return stored.StartsWith("sha256$"); }

    public static bool Verify(string stored, string password)
    {
        if (string.IsNullOrEmpty(stored) || string.IsNullOrEmpty(password)) return false;
        if (IsHashed(stored))
        {
            string[] p = stored.Split('$');
            return p.Length == 3 && Sha(p[1], password) == p[2];
        }
        return stored == password;   // plain password typed into Excel by the admin
    }

    static byte[] secret;
    static byte[] Secret()
    {
        if (secret != null) return secret;
        string path = Path.Combine(Store.Dir, "secret.key");
        if (File.Exists(path)) secret = Convert.FromBase64String(File.ReadAllText(path).Trim());
        else
        {
            secret = new byte[32];
            using (var rng = RandomNumberGenerator.Create()) rng.GetBytes(secret);
            File.WriteAllText(path, Convert.ToBase64String(secret));
        }
        return secret;
    }

    static string B64(byte[] b) { return Convert.ToBase64String(b).TrimEnd('=').Replace('+', '-').Replace('/', '_'); }
    static byte[] UnB64(string s)
    {
        s = s.Replace('-', '+').Replace('_', '/');
        while (s.Length % 4 != 0) s += "=";
        return Convert.FromBase64String(s);
    }
    static string Sign(string payload) { using (var h = new HMACSHA256(Secret())) return B64(h.ComputeHash(Encoding.UTF8.GetBytes(payload))); }

    public static string MakeToken(string userId, bool remember)
    {
        long exp = DateTime.UtcNow.AddHours(remember ? 24 * 14 : 12).Ticks;
        string payload = B64(Encoding.UTF8.GetBytes(userId + "|" + exp));
        return payload + "." + Sign(payload);
    }

    public static string ReadToken(string token)
    {
        try
        {
            if (string.IsNullOrEmpty(token)) return null;
            string[] p = token.Split('.');
            if (p.Length != 2 || Sign(p[0]) != p[1]) return null;
            string[] body = Encoding.UTF8.GetString(UnB64(p[0])).Split('|');
            if (body.Length != 2 || long.Parse(body[1]) < DateTime.UtcNow.Ticks) return null;
            return body[0];
        }
        catch { return null; }
    }
}

// ---------------------------------------------------------------- the API
public class NoteItApi : IHttpHandler
{
    public bool IsReusable { get { return false; } }
    static readonly JavaScriptSerializer Json = new JavaScriptSerializer { MaxJsonLength = int.MaxValue };

    Dictionary<string, object> body;
    Row me;

    string S(string key)
    {
        object v;
        if (body == null || !body.TryGetValue(key, out v) || v == null) return "";
        return Convert.ToString(v, CultureInfo.InvariantCulture).Trim();
    }
    bool Has(string key) { return body != null && body.ContainsKey(key); }
    bool B(string key) { object v; return body != null && body.TryGetValue(key, out v) && v is bool && (bool)v; }

    static bool Yes(string v) { v = (v ?? "").Trim().ToLowerInvariant(); return v == "yes" || v == "y" || v == "true" || v == "1" || v == "x"; }
    static bool Active(Row u) { string v = u.Get("Active").ToLowerInvariant(); return !(v == "no" || v == "n" || v == "false" || v == "0"); }
    static bool IsAdmin(Row u) { return Sec.Same(u.Get("Role"), "Admin"); }

    static Row FindUser(string id) { return Store.Users.Rows.FirstOrDefault(u => Sec.Same(u.Get("UserID"), id)); }
    static Row FindNote(string id) { return Store.Notes.Rows.FirstOrDefault(n => Sec.Same(n.Get("NoteID"), id)); }
    static Row FindTask(string id) { return Store.Tasks.Rows.FirstOrDefault(t => Sec.Same(t.Get("TaskID"), id)); }

    static List<string> Members(Row note)
    {
        return note.Get("Members").Split(new[] { ',', ';' }, StringSplitOptions.RemoveEmptyEntries)
            .Select(x => x.Trim()).Where(x => x != "").ToList();
    }
    static void SetMembers(Row note, IEnumerable<string> m) { note["Members"] = string.Join(", ", m.ToArray()); }

    string MyId { get { return me.Get("UserID"); } }
    bool IsOwner(Row note) { return Sec.Same(note.Get("CreatedBy"), MyId); }
    bool CanSeeNote(Row note) { return IsOwner(note) || Members(note).Any(m => Sec.Same(m, MyId)); }
    bool CanSeeTask(Row note, Row t)
    {
        string a = t.Get("AssignedTo");
        return a == "" || Sec.Same(a, MyId) || Sec.Same(t.Get("AddedBy"), MyId) || IsOwner(note);
    }
    bool InNote(Row note, string userId) { return Sec.Same(note.Get("CreatedBy"), userId) || Members(note).Any(m => Sec.Same(m, userId)); }

    Row NoteOrFail(string id)
    {
        Row n = FindNote(id);
        if (n == null || !CanSeeNote(n)) throw new ApiError("Note not found", 404);
        return n;
    }

    static object UserObj(Row u)
    {
        return new Dictionary<string, object> { { "id", u.Get("UserID") }, { "name", u.Get("Name") == "" ? u.Get("UserID") : u.Get("Name") }, { "role", IsAdmin(u) ? "Admin" : "User" } };
    }
    static object NoteObj(Row n)
    {
        return new Dictionary<string, object> { { "id", n.Get("NoteID") }, { "title", n.Get("Title") }, { "owner", n.Get("CreatedBy") }, { "members", Members(n) }, { "updated", n.Get("Updated") } };
    }
    static object TaskObj(Row t)
    {
        return new Dictionary<string, object> {
            { "id", t.Get("TaskID") }, { "noteId", t.Get("NoteID") }, { "text", t.Get("Task") },
            { "assignee", t.Get("AssignedTo") == "" ? null : t.Get("AssignedTo") },
            { "due", t.Get("DueDate") == "" ? null : t.Get("DueDate") }, { "done", Yes(t.Get("Done")) }, { "by", t.Get("AddedBy") } };
    }

    static string ValidDate(string d)
    {
        if (d == "") return "";
        DateTime dt;
        if (!DateTime.TryParseExact(d, "yyyy-MM-dd", CultureInfo.InvariantCulture, DateTimeStyles.None, out dt)) throw new ApiError("Invalid date");
        return dt.ToString("yyyy-MM-dd");
    }

    void Touch(Row note) { note["Updated"] = Store.Now(); }

    public void ProcessRequest(HttpContext ctx)
    {
        ctx.Response.ContentType = "application/json";
        ctx.Response.Cache.SetCacheability(HttpCacheability.NoCache);
        string raw;
        using (var r = new StreamReader(ctx.Request.InputStream, Encoding.UTF8)) raw = r.ReadToEnd();
        int status;
        string json = Handle(ctx.Request.QueryString["action"], ctx.Request.Headers["X-Auth-Token"], raw, out status);
        ctx.Response.StatusCode = status;
        ctx.Response.Write(json);
    }

    public static string Handle(string action, string token, string raw, out int status)
    {
        var api = new NoteItApi();
        object result;
        status = 200;
        try
        {
            api.body = string.IsNullOrWhiteSpace(raw) ? null : Json.DeserializeObject(raw) as Dictionary<string, object>;
            if (api.body == null) api.body = new Dictionary<string, object>();
            lock (Store.Lock)
            {
                Store.Load();
                result = api.Run((action ?? "").Trim(), token);
            }
        }
        catch (ApiError e) { status = e.Status; result = new Dictionary<string, object> { { "error", e.Message } }; }
        catch (IOException e) { Store.Users = null; Store.Notes = null; status = 409; result = new Dictionary<string, object> { { "error", "A data file is busy or open in Excel. Close it and try again. (" + e.Message + ")" } }; }
        catch (UnauthorizedAccessException) { status = 500; result = new Dictionary<string, object> { { "error", "The server cannot write to the App_Data folder. Give the IIS app pool 'Modify' permission on App_Data (see SETUP.txt)." } }; }
        catch (Exception e) { status = 500; result = new Dictionary<string, object> { { "error", "Server error: " + e.Message } }; }
        return Json.Serialize(result);
    }

    object Run(string action, string token)
    {
        if (action == "login") return Login();

        string uid = Sec.ReadToken(token);
        me = uid == null ? null : FindUser(uid);
        if (me == null || !Active(me)) throw new ApiError("Please log in again.", 401);

        switch (action)
        {
            case "me": return new Dictionary<string, object> { { "user", UserObj(me) } };
            case "changePassword": return ChangePassword();
            case "directory":
                return Store.Users.Rows.Where(Active).Select(UserObj).ToList();
            case "data": return Data();
            case "createNote": return CreateNote();
            case "updateNote": return UpdateNote();
            case "deleteNote": return DeleteNote();
            case "addMember": return AddMember();
            case "removeMember": return RemoveMember();
            case "leaveNote": return LeaveNote();
            case "addTask": return AddTask();
            case "updateTask": return UpdateTask();
            case "deleteTask": return DeleteTask();
            case "adminUsers": case "adminSaveUser": case "adminDeleteUser":
                if (!IsAdmin(me)) throw new ApiError("Only an admin can do this.", 403);
                if (action == "adminUsers") return AdminUsers();
                if (action == "adminSaveUser") return AdminSaveUser();
                return AdminDeleteUser();
        }
        throw new ApiError("Unknown action");
    }

    // ---- auth
    object Login()
    {
        Row u = FindUser(S("userId"));
        string pw = body.ContainsKey("password") ? Convert.ToString(body["password"]) : "";
        if (u == null || !Sec.Verify(u.Get("Password"), pw)) throw new ApiError("Incorrect User ID or password.", 401);
        if (!Active(u)) throw new ApiError("This User ID is not active. Contact your admin.", 403);
        if (!Sec.IsHashed(u.Get("Password")))
        {
            u["Password"] = Sec.Hash(pw);           // protect a plain password typed into Excel
            try { Store.SaveUsers(); } catch (ApiError) { }  // fine if Excel has the file open
        }
        bool mustChange = Sec.Same(u.Get("UserID"), "admin") && pw == "admin123";
        return new Dictionary<string, object> { { "token", Sec.MakeToken(u.Get("UserID"), B("remember")) }, { "user", UserObj(u) }, { "mustChange", mustChange } };
    }

    object ChangePassword()
    {
        string oldPw = Has("oldPassword") ? Convert.ToString(body["oldPassword"]) : "";
        string newPw = Has("newPassword") ? Convert.ToString(body["newPassword"]) : "";
        if (!Sec.Verify(me.Get("Password"), oldPw)) throw new ApiError("Current password is wrong.");
        if (newPw.Length < 6) throw new ApiError("New password must be at least 6 characters.");
        me["Password"] = Sec.Hash(newPw);
        Store.SaveUsers();
        return new Dictionary<string, object> { { "ok", true } };
    }

    // ---- notes & tasks
    object Data()
    {
        var notes = Store.Notes.Rows.Where(CanSeeNote).ToList();
        var ids = new HashSet<string>(notes.Select(n => n.Get("NoteID")), StringComparer.OrdinalIgnoreCase);
        var tasks = Store.Tasks.Rows.Where(t => ids.Contains(t.Get("NoteID")))
            .Where(t => CanSeeTask(notes.First(n => Sec.Same(n.Get("NoteID"), t.Get("NoteID"))), t)).ToList();
        return new Dictionary<string, object> { { "notes", notes.Select(NoteObj).ToList() }, { "tasks", tasks.Select(TaskObj).ToList() } };
    }

    object CreateNote()
    {
        var n = new Row();
        n["NoteID"] = Store.NextId(Store.Notes, "NoteID", "N");
        n["Title"] = S("title");
        n["CreatedBy"] = MyId; n["Members"] = ""; n["Created"] = Store.Now(); n["Updated"] = Store.Now();
        Store.Notes.Rows.Add(n);
        Store.SaveTasks();
        return NoteObj(n);
    }

    object UpdateNote()
    {
        Row n = NoteOrFail(S("id"));
        if (!IsOwner(n)) throw new ApiError("Only the creator can rename this note.", 403);
        n["Title"] = S("title"); Touch(n);
        Store.SaveTasks();
        return NoteObj(n);
    }

    object DeleteNote()
    {
        Row n = NoteOrFail(S("id"));
        if (!IsOwner(n)) throw new ApiError("Only the creator can delete this note.", 403);
        Store.Tasks.Rows.RemoveAll(t => Sec.Same(t.Get("NoteID"), n.Get("NoteID")));
        Store.Notes.Rows.Remove(n);
        Store.SaveTasks();
        return new Dictionary<string, object> { { "ok", true } };
    }

    object AddMember()
    {
        Row n = NoteOrFail(S("noteId"));
        if (!IsOwner(n)) throw new ApiError("Only the creator can add people.", 403);
        Row u = FindUser(S("userId"));
        if (u == null || !Active(u)) throw new ApiError("No active user with ID \"" + S("userId") + "\".");
        if (InNote(n, u.Get("UserID"))) throw new ApiError("Already on this note.");
        var m = Members(n); m.Add(u.Get("UserID")); SetMembers(n, m); Touch(n);
        Store.SaveTasks();
        return NoteObj(n);
    }

    void DropFromNote(Row n, string userId)
    {
        SetMembers(n, Members(n).Where(x => !Sec.Same(x, userId)));
        foreach (Row t in Store.Tasks.Rows)
            if (Sec.Same(t.Get("NoteID"), n.Get("NoteID")) && Sec.Same(t.Get("AssignedTo"), userId)) t["AssignedTo"] = "";
        Touch(n);
    }

    object RemoveMember()
    {
        Row n = NoteOrFail(S("noteId"));
        if (!IsOwner(n)) throw new ApiError("Only the creator can remove people.", 403);
        DropFromNote(n, S("userId"));
        Store.SaveTasks();
        return NoteObj(n);
    }

    object LeaveNote()
    {
        Row n = NoteOrFail(S("noteId"));
        if (IsOwner(n)) throw new ApiError("The creator cannot leave their own note.");
        DropFromNote(n, MyId);
        Store.SaveTasks();
        return new Dictionary<string, object> { { "ok", true } };
    }

    string CheckAssignee(Row n, string a)
    {
        if (a == "") return "";
        if (!InNote(n, a)) throw new ApiError("That person is not on this note.");
        Row u = FindUser(a);
        return u != null ? u.Get("UserID") : a;
    }

    object AddTask()
    {
        Row n = NoteOrFail(S("noteId"));
        if (S("text") == "") throw new ApiError("Task text is empty.");
        var t = new Row();
        t["TaskID"] = Store.NextId(Store.Tasks, "TaskID", "T");
        t["NoteID"] = n.Get("NoteID"); t["Task"] = S("text");
        t["AssignedTo"] = CheckAssignee(n, S("assignee"));
        t["DueDate"] = ValidDate(S("due")); t["Done"] = "No";
        t["AddedBy"] = MyId; t["Created"] = Store.Now();
        Store.Tasks.Rows.Add(t); Touch(n);
        Store.SaveTasks();
        return TaskObj(t);
    }

    Row TaskOrFail(out Row note)
    {
        Row t = FindTask(S("id"));
        note = t == null ? null : FindNote(t.Get("NoteID"));
        if (t == null || note == null || !CanSeeNote(note) || !CanSeeTask(note, t)) throw new ApiError("Task not found", 404);
        return t;
    }

    object UpdateTask()
    {
        Row n;
        Row t = TaskOrFail(out n);
        bool canEdit = IsOwner(n) || Sec.Same(t.Get("AddedBy"), MyId);
        if (Has("done")) t["Done"] = B("done") ? "Yes" : "No";
        if (Has("assignee") || Has("due") || Has("text"))
        {
            if (!canEdit) throw new ApiError("Only the note creator or the person who added the task can change it.", 403);
            if (Has("assignee")) t["AssignedTo"] = CheckAssignee(n, S("assignee"));
            if (Has("due")) t["DueDate"] = ValidDate(S("due"));
            if (Has("text") && S("text") != "") t["Task"] = S("text");
        }
        Touch(n);
        Store.SaveTasks();
        return TaskObj(t);
    }

    object DeleteTask()
    {
        Row n;
        Row t = TaskOrFail(out n);
        if (!(IsOwner(n) || Sec.Same(t.Get("AddedBy"), MyId))) throw new ApiError("Only the note creator or the person who added the task can delete it.", 403);
        Store.Tasks.Rows.Remove(t); Touch(n);
        Store.SaveTasks();
        return new Dictionary<string, object> { { "ok", true } };
    }

    // ---- admin
    object AdminUsers()
    {
        return Store.Users.Rows.Select(u => new Dictionary<string, object> {
            { "id", u.Get("UserID") }, { "name", u.Get("Name") }, { "role", IsAdmin(u) ? "Admin" : "User" },
            { "active", Active(u) }, { "created", u.Get("Created") },
            { "notes", Store.Notes.Rows.Count(n => Sec.Same(n.Get("CreatedBy"), u.Get("UserID"))) } }).ToList();
    }

    int ActiveAdmins(Row except) { return Store.Users.Rows.Count(u => u != except && IsAdmin(u) && Active(u)); }

    object AdminSaveUser()
    {
        string id = S("userId");
        bool isNew = B("isNew");
        string pw = Has("password") ? Convert.ToString(body["password"]) : "";
        string role = Sec.Same(S("role"), "Admin") ? "Admin" : "User";
        bool active = !Has("active") || B("active");
        Row u = FindUser(id);
        if (isNew)
        {
            if (!Regex.IsMatch(id, "^[A-Za-z0-9_.-]{2,30}$")) throw new ApiError("User ID must be 2-30 characters: letters, numbers, _ . or -");
            if (u != null) throw new ApiError("User ID \"" + id + "\" already exists.");
            if (pw.Length < 6) throw new ApiError("Password must be at least 6 characters.");
            u = new Row();
            u["UserID"] = id; u["Created"] = Store.Now();
            Store.Users.Rows.Add(u);
        }
        else
        {
            if (u == null) throw new ApiError("User not found", 404);
            if ((role != "Admin" || !active) && IsAdmin(u) && ActiveAdmins(u) == 0) throw new ApiError("There must be at least one active admin.");
            if (pw != "" && pw.Length < 6) throw new ApiError("Password must be at least 6 characters.");
        }
        u["Name"] = S("name") == "" ? id : S("name");
        u["Role"] = role;
        u["Active"] = active ? "Yes" : "No";
        if (pw != "") u["Password"] = Sec.Hash(pw);
        Store.SaveUsers();
        return new Dictionary<string, object> { { "ok", true } };
    }

    object AdminDeleteUser()
    {
        Row u = FindUser(S("userId"));
        if (u == null) throw new ApiError("User not found", 404);
        if (u == me) throw new ApiError("You cannot delete your own User ID.");
        if (IsAdmin(u) && ActiveAdmins(u) == 0) throw new ApiError("There must be at least one active admin.");
        int owned = Store.Notes.Rows.Count(n => Sec.Same(n.Get("CreatedBy"), u.Get("UserID")));
        if (owned > 0) throw new ApiError("This user created " + owned + " note(s). Set them to Inactive instead, so their notes stay available.");
        bool changed = false;
        foreach (Row n in Store.Notes.Rows.Where(n => InNote(n, u.Get("UserID"))).ToList()) { DropFromNote(n, u.Get("UserID")); changed = true; }
        Store.Users.Rows.Remove(u);
        Store.SaveUsers();
        if (changed) Store.SaveTasks();
        return new Dictionary<string, object> { { "ok", true } };
    }
}
