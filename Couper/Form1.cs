﻿using Microsoft.Office.Interop.Outlook;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.Linq;
using System.Reflection;
using System.Threading.Tasks;
using System.Windows.Forms;
using Application = Microsoft.Office.Interop.Outlook.Application;
using Exception = System.Exception;
using Action = System.Action;
using Folder = Microsoft.Office.Interop.Outlook.Folder;
using Microsoft.Office.Interop.OneNote;
using System.Xml.Linq;
using System.IO;
using System.Text;
using System.Xml.Serialization;
using BrightIdeasSoftware;
using System.Runtime.InteropServices;

namespace Couper
{
    public partial class Form1 : Form
    {
        private string _pageId;
        private string _usedFile;
        private Details[] _details;
        private bool _allSelected;
        private Settings _settings;
        private string _settingsFile;
        private PropertyInfo[] _columns;
        private Application _outlookApplication;
        private Microsoft.Office.Interop.OneNote.Application _app;

        private const string TitleCode = "קוד שובר";
        private const string TitleCode2 = "שובר לרכישה";
        private const string TitleAmount = "סכום ההזמנה";
        private const string TitleAmount2 = "החיוב בסיבוס שלך";
        private const string TitleExpires = "תוקף";
        private const string TitleLocation = "סניף";
        private const string TitleLocation2 = "קיבלנו את הזמנת השובר שלך";
        private const string TitleDate = "תאריך";
        private const string TitleUsed = "משומש";
        private const string TitleLink = "ללחוץ כאן";

        private const string Subject = "שובר על סך";

        private const string PageName = "Couper";
        private const string SectionName = "Shopping";
        private const string Sum = "Sum";
        private const string DateFormat = "dd/MM/yyyy";


        public Form1()
        {
            InitializeComponent();

            _usedFile = Path.Combine(System.Windows.Forms.Application.LocalUserAppDataPath, "used.txt");
            _settingsFile = Path.Combine(System.Windows.Forms.Application.LocalUserAppDataPath, "settings.ini");
        }

        private async void Form1_Load(object sender, EventArgs e)
        {
            try
            {

                Text = "Couper " + Properties.Settings.Default.Version;

                LoadSettings();

                EnableButton(tsOneNote, false);

                _columns = new Details().GetType().GetProperties();

                lstResults.UseHyperlinks = true;
                lstResults.HyperlinkClicked += LstResults_HyperlinkClicked;

                Generator.GenerateColumns(lstResults, typeof(Details), true);
                lstResults.AutoResizeColumns();

                lstResults.FormatRow += FormatRow;

                var menu = new ContextMenuStrip();
                menu.Items.Add(new ToolStripMenuItem("Copy", Properties.Resources.Copy, (s, _) => OnCopy()));

                lstResults.ContextMenuStrip = menu;

                if (!File.Exists(_settingsFile))
                {
                    await Task.Delay(1000);
                    MessageBox.Show(this, MailMessage() + "\n\n" + SyncMessage(), "Couper", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (Exception ex)
            {
                Log(ex);
            }
        }

        private void LstResults_HyperlinkClicked(object sender, HyperlinkClickedEventArgs e)
        {
            RunSafe(() =>
            {
                var detail = (Details)e.Item.RowObject;
                e.Url = detail.Link;
            });
        }

        private static void FormatRow(object sender, FormatRowEventArgs e)
        {
            var msg = (Details)e.Model;

            if (msg.Used)
            {
                e.Item.ForeColor = Color.DarkSlateGray;
                return;
            }

            if (msg.Expires < DateTime.Now || msg.Expires - DateTime.Now < TimeSpan.FromDays(3))
            {
                e.Item.ForeColor = Color.DarkRed;
                return;
            }

            e.Item.ForeColor = Color.DarkBlue;

        }

        private void OnCopy()
        {
            RunSafe(() =>
            {
                var lines = lstResults.CheckedObjectsEnumerable.Cast<Details>()
                    .Select(d => string.Join("\t", _columns.Select(f => f.GetValue(d))));

                Clipboard.SetText(string.Join("\r\n", lines));
            });
        }

        private string GetField(string body, string name)
        {
            if (!body.Contains(name))
            {
                return null;
            }

            return body.Split(new[] { name }, StringSplitOptions.None)[1].Trim().Split('\r')[0].Trim();
        }

        private string GetField(string body, string name, string name2)
        {
            var res = GetField(body, name);
            try
            {
                if (res != null)
                {
                    return res;
                }

                if (name2 != null)
                {
                    res = GetField(body, name2);
                    if (res != null)
                    {
                        return res;
                    }
                }

                return null;
            }
            catch (Exception ex)
            {
                Log(ex);
                return null;
            }
        }

        private int IndexOf(string col)
        {
            return lstResults.Columns.Cast<OLVColumn>().First(c => c.Text == col).Index;
        }

        private bool ShowQuestion(string title, string msg)
        {
            return MessageBox.Show(this, msg + ".\nAre you sure?", title, MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes;
        }

        private async void btnGo_Click(object sender, EventArgs e)
        {
            try
            {
                EnableButton(btnGo, false);

                _allSelected = true;

                SetProg(true);

                if (string.IsNullOrEmpty(tsNotebook.Text))
                {
                    if (!ShowQuestion("OneNote Notebook",
                        "Running without a notebook will not sync to OneNote"))
                    {
                        return;
                    }
                }

                if (string.IsNullOrEmpty(txtFolder.Text))
                {
                    if (!ShowQuestion("Cibus Folder",
                        "Running without a given folder (such as Cibus) might take some time"))
                    {
                        return;
                    }

                    txtFolder.Text = "Inbox";
                }

                var folder = txtFolder.Text;

                SaveSettings();

                var days = Convert.ToInt32(txtDays.Text);

                await Task.Run(() => GetItems(days, txtFolder.Text));
            }
            catch (Exception ex)
            {
                Log(ex);
            }
        }

        private void UpdateItems(Details[] detailsFromNote, Details[] details)
        {
            RunSafe(() =>
            {
                if (InvokeRequired)
                {
                    BeginInvoke((Action)(() => UpdateItems(detailsFromNote, details)));
                    return;
                }

                if (detailsFromNote != null)
                {
                    foreach (var detail in details)
                    {
                        var exist = detailsFromNote.FirstOrDefault(d => d.Number == detail.Number);
                        if (exist != null)
                        {
                            detail.Used = exist.Used;
                        }
                    }
                }

                _details = details;

                UpdateInGui(details);

                UpdateSum();

                var used = LoadUsed();
                used = used.Concat(details.Where(d => d.Used).Select(d => d.Number)).ToArray();

                File.WriteAllLines(_usedFile, used);
            });
        }

        private void UpdateInGui(Details[] details)
        {
            lstResults.Items.Clear();

            lstResults.AddObjects(details);
            lstResults.CheckAll();

            lstResults.AutoResizeColumns();
        }

        private void Log(Exception ex)
        {
            Log(ex.Message, true);
        }

        private void Log(string message, bool error = false)
        {
            try
            {
                if (InvokeRequired)
                {
                    BeginInvoke((Action)(() => Log(message, error)));
                    return;
                }

                var time = DateTime.Now.ToString("HH:mm:ss");

                var item = new ListViewItem(time);
                item.SubItems.Add(message);

                if (error)
                {
                    item.ForeColor = Color.DarkRed;
                }

                lstLog.Items.Add(item);

                lstLog.Items[lstLog.Items.Count - 1].EnsureVisible();
            }
            catch
            {
                // ignored
            }
        }

        private string[] LoadUsed()
        {
            if (!File.Exists(_usedFile))
            {
                return new string[0];
            }

            return File.ReadAllLines(_usedFile);
        }

        private async void Application_AdvancedSearchComplete(Search search, int days)
        {
            try
            {
                var now = DateTime.Now;
                var items = new List<MailItem>();

                foreach (var result in search.Results)
                {
                    if (!(result is MailItem))
                    {
                        continue;
                    }

                    var item = (MailItem)result;

                    if (now - item.ReceivedTime > TimeSpan.FromDays(days))
                    {
                        continue;
                    }

                    var subject = item.Subject;
                    if (!subject.Contains(Subject))
                    {
                        continue;
                    }

                    if (item.Sender?.Name?.Contains("Cibus") == true)
                    {
                        items.Add(item);
                        Log($"{item.Subject} ({item.ReceivedTime})");
                    }
                }

                var detailsFromNote = SyncToOneNote(null, false);

                UpdateSum();

                var details = items
                    .Select(i => new Details
                    {
                        Number = GetField(i.Body, TitleCode + ":", TitleCode2),
                        Amount = Convert.ToInt32(GetField(i.Body, TitleAmount + ":", TitleAmount2 + ":").Split(' ')[0].Replace("₪", "").Split('.')[0]),
                        Expires = ParseDate(GetField(i.Body, TitleExpires + " ", null) ?? i.ReceivedTime.ToString(DateFormat)),
                        Location = GetField(i.Body, TitleLocation + ":", TitleLocation2),
                        Date = ParseDate(GetField(i.Body, TitleDate + ":", null) ?? i.ReceivedTime.ToString(DateFormat)),
                        Link = GetField(i.Body, TitleLink, null).Replace("<", "").Replace(">", "")
                    })
               .Distinct()
               .OrderBy(i => i.Date)
               .ThenByDescending(i => Convert.ToInt32(i.Amount))
               .ToArray();

                var used = LoadUsed();
                if (used.Length > 0)
                {
                    foreach (var detail in details)
                    {
                        detail.Used = used.Contains(detail.Number);
                    }
                }

                UpdateItems(detailsFromNote, details);

                Log($"Found {details.Length} items in mail");

                if (details.Length > 0)
                {
                    await Task.Run(() => SyncToOneNote(details, true));
                }
            }
            catch (Exception ex)
            {
                Log(ex);
            }
            finally
            {
                Log("Done.");

                EnableButton(btnGo, true);
                SetProg(false);
            }
        }

        private void GetItems(int days, string cibusFolder)
        {
            _outlookApplication = new Application();
            _outlookApplication.AdvancedSearchComplete += (s) => Application_AdvancedSearchComplete(s, days);

            NameSpace outlookNamespace = _outlookApplication.GetNamespace("MAPI");
            MAPIFolder inboxFolder = outlookNamespace.GetDefaultFolder(OlDefaultFolders.olFolderInbox);
            var items = new List<MailItem>();

            Log($"Fetching mails from the last {days} days (Folder: {cibusFolder})");

            string scope = null;
            var dateStart = DateTime.Now.AddDays(-1 * days);

            string filter = $"urn:schemas:mailheader:subject LIKE \'%{Subject}%\' AND urn:schemas:httpmail:datereceived > '{dateStart}'";

            Search advancedSearch = null;
            NameSpace ns = null;

            Folder folder;

            if (cibusFolder == "Inbox")
            {
                folder = (Folder)inboxFolder;
            }
            else
            {
                folder = inboxFolder.Folders.Cast<Folder>().FirstOrDefault(f => f.Name == cibusFolder);
            }

            try
            {
                ns = _outlookApplication.GetNamespace("MAPI");

                scope = "\'" + folder.FolderPath + "\'";
                advancedSearch = _outlookApplication.AdvancedSearch(scope, filter, false, "Couper Search");
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "An eexception is thrown");
            }
            finally
            {
                if (advancedSearch != null)
                {
                    Marshal.ReleaseComObject(advancedSearch);
                }
                if (inboxFolder != null)
                {
                    Marshal.ReleaseComObject(inboxFolder);
                }
                if (ns != null)
                {
                    Marshal.ReleaseComObject(ns);
                }
            }
        }

        private static DateTime ParseDate(string date)
        {
            if (DateTime.TryParseExact(date, DateFormat, CultureInfo.CurrentCulture, DateTimeStyles.None, out var result))
            {
                return result;
            }

            if (DateTime.TryParseExact(date, "dd/M/yyyy", CultureInfo.CurrentCulture, DateTimeStyles.None, out result))
            {
                return result;
            }

            return DateTime.Parse(date);
        }

        private void RunSafe(Action action)
        {
            try
            {
                action.Invoke();
            }
            catch (Exception ex)
            {
                Log(ex);
            }
        }

        private async Task RunSafeAsync(Action action)
        {
            await Task.Run(() => RunSafe(action));
        }

        private void SaveSettings()
        {
            RunSafe(() =>
            {
                _settings = new Settings
                {
                    Notebook = tsNotebook.Text,
                    CibusFolder = txtFolder.Text
                };

                using (StreamWriter writer = new StreamWriter(_settingsFile, false, Encoding.Unicode))
                {
                    var serializer = new XmlSerializer(typeof(Settings));
                    serializer.Serialize(writer, _settings);
                    writer.Flush();
                }
            });
        }

        private void LoadSettings()
        {
            RunSafe(() =>
            {
                if (!File.Exists(_settingsFile))
                {
                    return;
                }

                using (var stream = new StreamReader(_settingsFile, Encoding.Unicode))
                {
                    var serializer = new XmlSerializer(typeof(Settings));
                    _settings = (Settings)serializer.Deserialize(stream);
                }

                tsNotebook.Text = _settings.Notebook;

                if (!string.IsNullOrEmpty(_settings.CibusFolder))
                {
                    txtFolder.Text = _settings.CibusFolder;
                }
            });
        }

        private void Form1_KeyDown(object sender, KeyEventArgs e)
        {
            RunSafe(() =>
            {
                if (e.KeyCode == Keys.F5 || e.KeyCode == Keys.Enter)
                {
                    btnGo.PerformClick();
                }
            });
        }

        private void UpdateSum()
        {
            RunSafe(() =>
            {
                if (InvokeRequired)
                {
                    BeginInvoke((Action)(() => UpdateStyles()));
                    return;
                }

                lblSum.Text = lstResults.CheckedObjectsEnumerable.Cast<Details>()
               .Where(i => !i.Used)
               .Sum(i => Convert.ToInt32(i.Amount)).ToString();
            });
        }

        private void lstResults_ItemChecked(object sender, System.Windows.Forms.ItemCheckedEventArgs e)
        {
            RunSafe(() =>
            {
                UpdateSum();
            });
        }

        private void btnSelectAll_Click(object sender, EventArgs e)
        {
            RunSafe(() =>
            {
                _allSelected = !_allSelected;

                if (_allSelected)
                {
                    lstResults.CheckAll();
                }
                else
                {
                    lstResults.UncheckAll();
                }
            });
        }

        private void btnCopy_Click(object sender, EventArgs e)
        {
            RunSafe(() =>
            {
                OnCopy();
            });
        }

        private Details ParseRow(XNamespace ns, XElement row)
        {
            var cells = row.Descendants(ns + "Cell")
                    .Select(c => (GetCellValue(c), IsComplete(ns, c)))
                    .ToArray();

            return new Details
            {
                Amount = Convert.ToInt32(cells[1].Item1.Item1),
                Number = cells[2].Item1.Item1,
                Date = ParseDate(cells[3].Item1.Item1),
                Location = cells[4].Item1.Item1,
                Expires = ParseDate(cells[5].Item1.Item1),
                Link = cells[2].Item1.Item2,
                Used = cells.Any(c => c.Item2)
            };
        }

        private Details[] ParseDetails(XNamespace ns, XElement outline)
        {
            // skip header
            var rows = outline.Descendants(ns + "Row").Skip(1);

            var details = new List<Details>();

            foreach (var row in rows)
            {
                var cells = row.Descendants(ns + "Cell")
                    .Select(c => (GetCellValue(c).Item1, IsComplete(ns, c))).ToArray();

                details.Add(ParseRow(ns, row));
            }

            return details.ToArray();
        }

        private (string, string) GetCellValue(XElement cell)
        {
            if (!cell.Value.Contains(">"))
            {
                return (cell.Value, null);
            }

            var data = cell.Value.Split('>')[1].Split('<')[0];

            if (cell.Value.Contains("href"))
            {
                return (data, cell.Value.Split('\"')[1]);
            }

            // change "<span\nstyle='direction:ltr;unicode-bidi:embed' lang=en-US>200</span>" to 200
            return (data, null);
        }

        private bool IsComplete(XNamespace ns, XElement cell)
        {
            var tag = cell.Descendants(ns + "Tag").FirstOrDefault();
            if (tag == null)
            {
                return false;
            }

            return tag.Attribute("completed").Value == "true";
        }

        private void EnableButton(ToolStripButton btn, bool enabled)
        {
            RunSafe(() =>
            {
                if (InvokeRequired)
                {
                    BeginInvoke((Action)(() => EnableButton(btn, enabled)));
                    return;
                }

                btn.Enabled = enabled;
            });
        }

        private Details[] SyncToOneNote(Details[] details, bool update)
        {
            try
            {
                if (string.IsNullOrEmpty(_settings.Notebook))
                {
                    Log("Not syncing since notebook name is empty");
                    return null;
                }

                Log($"Syncing with OneNote ({_settings.Notebook}\\{SectionName})");

                _app = new Microsoft.Office.Interop.OneNote.Application();
                _app.GetHierarchy(null, HierarchyScope.hsPages, out string notebookXml);

                var mainDoc = XDocument.Parse(notebookXml);
                var ns = mainDoc.Root.Name.Namespace;

                var notebook = mainDoc.Descendants(ns + "Notebook").Where(n => n.Attribute("name").Value == _settings.Notebook).FirstOrDefault();
                if (notebook == null)
                {
                    throw new Exception("Failed to find notebook " + _settings.Notebook);
                }

                var section = notebook.Descendants(ns + "Section").Where(n => n.Attribute("name").Value == SectionName).FirstOrDefault();
                if (section == null)
                {
                    throw new Exception($"Failed to find section {SectionName}. Please create it under {_settings.Notebook}");
                }

                string sectionId = section.Attribute("ID").Value;

                var pageNode = section.Descendants(ns + "Page").Where(n => n.Attribute("name").Value == PageName).LastOrDefault();
                string xml;

                if (pageNode == null)
                {
                    if (!update)
                    {
                        return null;
                    }

                    xml = CreateNewPage(PageName, ns, sectionId);
                }
                else
                {
                    _pageId = pageNode.Attribute("ID").Value;
                    _app.GetPageContent(_pageId, out xml, PageInfo.piAll);
                }

                EnableButton(tsOneNote, true);

                var doc = XDocument.Parse(xml);
                var outline = doc.Descendants(ns + "Outline").FirstOrDefault();

                var table = outline.Descendants(ns + "Table").First();
                var content = outline.ToString();

                var existingDetails = ParseDetails(ns, outline);

                if (!update)
                {
                    return existingDetails;
                }

                if (existingDetails.Length == 0)
                {
                    existingDetails = details;
                }

                var rows = table.Descendants(ns + "Row").ToList();
                table.Descendants(ns + "Row").Remove();

                foreach (var detail in details)
                {
                    var existing = rows.FirstOrDefault(r => r.Descendants(ns + "Cell").Any(c => c.Value.Contains(detail.Number)));
                    if (existing != null)
                    {
                        var parsed = ParseRow(ns, existing);
                        if (parsed.Used == detail.Used)
                        {
                            continue;
                        }

                        rows.Remove(existing);
                    }

                    rows.Add(
                    new XElement(ns + "Row",
                        BuildCell(ns, "", true, detail.Used),
                        BuildCell(ns, detail.Amount.ToString()),
                        BuildCell(ns, detail.Number, link: detail.Link),
                        BuildCell(ns, detail.Date.ToString(DateFormat)),
                        BuildCell(ns, detail.Location),
                        BuildCell(ns, detail.Expires.ToString(DateFormat))
                        ));
                }

                // header line
                table.Add(rows[0]);

                foreach (var row in rows.Skip(1).OrderBy(r => ParseRow(ns, r).Date))
                {
                    if(ParseRow(ns, row).Used)
                    {
                        continue;
                    }

                    table.Add(row);
                }

                var sumElem = outline.Descendants(ns + "T").Where(e => e.Value.Contains(Sum)).First();
                var sum = existingDetails
                    .Where(i => !i.Used)
                    .Sum(i => Convert.ToInt32(i.Amount));

                sumElem.Value = $"{Sum}: {sum}";

                _app.UpdatePageContent(doc.ToString());

                return existingDetails;
            }
            catch (Exception ex)
            {
                Log(ex);
                return null;
            }
            finally
            {
                Log("Sync Done.");
            }
        }

        private async void tsOneNote_Click(object sender, EventArgs e)
        {
            await RunSafeAsync(() =>
            {
                Log("Opening OneNote..");

                if (string.IsNullOrEmpty(_pageId) || _app == null)
                {
                    throw new Exception("OneNote not inistialized yet..");
                }

                _app = new Microsoft.Office.Interop.OneNote.Application();
                _app.GetHyperlinkToObject(_pageId, "", out var link);

                _app.NavigateToUrl(link);
            });
        }

        private void SetProg(bool busy)
        {
            RunSafe(() =>
            {
                if (InvokeRequired)
                {
                    BeginInvoke((Action)(() => SetProg(busy)));
                    return;
                }

                tsProg.Style = busy ? ProgressBarStyle.Marquee : ProgressBarStyle.Continuous;
            });
        }


        private string CreateNewPage(string pageName,
            XNamespace ns,
            string sectionId)
        {
            Log($"Creating new OneNote page ({pageName})");

            _app.CreateNewPage(sectionId, out _pageId, NewPageStyle.npsBlankPageWithTitle);

            XElement newPage = new XElement(ns + "Page");
            newPage.SetAttributeValue("ID", _pageId);
            newPage.SetAttributeValue("name", pageName);

            var tag = new XElement(ns + "TagDef",
                       new XAttribute("index", "0"),
                       new XAttribute("type", "0"),
                       new XAttribute("symbol", "3"),
                       new XAttribute("name", "To Do"));

            newPage.Add(tag);

            newPage.Add(new XElement(ns + "Title",
                            new XElement(ns + "OE",
                                new XElement(ns + "T",
                                    new XCData(pageName)))));

            var outline = new XElement(ns + "Outline");

            var columns = new List<XElement>();

            for (int i = 0; i < 6; ++i)
            {
                columns.Add(new XElement(ns + "Column",
                  new XAttribute("index", $"{i}"),
                  new XAttribute("width", "120")));
            }

            var row = new XElement(ns + "Row",
                BuildCell(ns, ""),
                BuildCell(ns, TitleAmount),
                BuildCell(ns, TitleCode),
                BuildCell(ns, TitleDate),
                BuildCell(ns, TitleLocation),
                BuildCell(ns, TitleExpires));

            var table = new XElement(ns + "Table",
                  new XAttribute("bordersVisible", "true"),
                  new XAttribute("hasHeaderRow", "true"),
                  new XElement(ns + "Columns",
                  columns),
                  row);

            var sum = new XElement(ns + "OE",
                new XElement(ns + "T",
                new XCData($"{Sum}: 0\n\n")));

            outline.Add(new XElement(ns + "OEChildren",
                        sum,
                        new XElement(ns + "OE",
                            new XElement(ns + "T",
                                new XCData($""))),
                        new XElement(ns + "OE",
                            table)));

            newPage.Add(outline);

            return newPage.ToString();
        }

        private XElement BuildCell(XNamespace ns, string data, bool addCheck = false, bool isChecked = false, string link = null)
        {
            var tag = new XElement(ns + "Tag",
                                new XAttribute("index", "0"),
                                new XAttribute("completed", isChecked ? "true" : "false"));

            if (!string.IsNullOrEmpty(link))
            {
                data = $"<a href=\"{link}\">{data}</a>";
            }

            var child = new XElement(ns + "OE",
                                new XElement(ns + "T",
                                    new XCData(data)));

            if (addCheck)
            {
                child.AddFirst(tag);
            }

            return new XElement(ns + "Cell",
                    new XElement(ns + "OEChildren",
                        child));
        }

        private async void tsSync_Click(object sender, EventArgs e)
        {
            EnableButton(tsSync, false);

            SetProg(true);

            var items = lstResults.CheckedObjectsEnumerable.Cast<Details>().ToArray();

            await RunSafeAsync(() =>
            {
                var detailsFromNote = SyncToOneNote(items, true);
                UpdateItems(detailsFromNote, detailsFromNote);

            });

            EnableButton(tsSync, true);

            SetProg(false);
        }

        private string SyncMessage()
        {
            return "In order to sync with OneNote please do the following:\n\n" +
                "* Enter the notebook name\n" +
                $"* Create a section called {SectionName}\n";
        }

        private string MailMessage()
        {
            return "Supply the mail folder where the Cibus mails go to.\n" +
                "If you do not have such folder it is recommended to create one (such as Cibus)\n" +
                "Otherwise the Inbox folder will be searched (which may take some time)";
        }

        private void lnkOneNote_Click(object sender, EventArgs e)
        {
            MessageBox.Show(this, SyncMessage(), "Cibus Folder", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void lnkFolder_Click(object sender, EventArgs e)
        {
            MessageBox.Show(this, MailMessage(), "Cibus Folder", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void chkHideUsed_CheckedChanged(object sender, EventArgs e)
        {
            RunSafe(() =>
            {
                if (_details == null || _details.All(d => !d.Used))
                {
                    return;
                }

                var details = _details;

                if (chkHideUsed.Checked == true)
                {
                    details = _details.Where(d => !d.Used).ToArray();
                }

                UpdateInGui(details);
            });
        }
    }

    public class Settings
    {
        public string Notebook;
        public string CibusFolder;
    }

    public class Details
    {
        [OLVColumn(AspectToStringFormat = "{0:d}")]
        public DateTime Date { get; set; }
        public int Amount { get; set; }
        [OLVColumn(Hyperlink = true)]
        public string Number { get; set; }
        [OLVColumn(AspectToStringFormat = "{0:d}")]
        public DateTime Expires { get; set; }
        public string Location { get; set; }
        public bool Used { get; set; }
        public string Link;

        public override bool Equals(object obj)
        {
            if (obj is Details details)
            {
                return Number == details.Number;
            }

            return false;
        }

        public override int GetHashCode()
        {
            return base.GetHashCode();
        }
    }
}
