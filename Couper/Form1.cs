using Microsoft.Office.Interop.Outlook;
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

namespace Couper
{
    public partial class Form1 : Form
    {
        private bool _allSelected;
        private Settings _settings;
        private string _settingsFile;
        private FieldInfo[] _columns;

        private const string TitleCode = "קוד שובר";
        private const string TitleAmount = "סכום ההזמנה";
        private const string TitleExpires = "תוקף";
        private const string TitleLocation = "סניף";
        private const string TitleDate = "תאריך";
        private const string TitleUsed = "משומש";

        private const string PageName = "Couper";
        private const string SectionName = "Shopping";

        public Form1()
        {
            InitializeComponent();

            _settingsFile = Path.Combine(System.Windows.Forms.Application.LocalUserAppDataPath, "settings.ini");
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            RunSafe(() =>
            {
                LoadSettings();

                _columns = new Details().GetType().GetFields();

                lstResults.Columns.AddRange(_columns.Select(f => new ColumnHeader
                {
                    Text = f.Name,
                    Width = 150
                }).ToArray());

                var menu = new ContextMenuStrip();
                menu.Items.Add(new ToolStripMenuItem("Copy", Properties.Resources.Copy, (s, _) => OnCopy()));

                lstResults.ContextMenuStrip = menu;
            });
        }

        private void OnCopy()
        {
            RunSafe(() =>
            {
                var lines = lstResults.CheckedItems.Cast<ListViewItem>()
                    .Select(i => (Details)i.Tag)
                    .Select(d => string.Join("\t", d.GetType().GetFields().Select(f => f.GetValue(d))));

                Clipboard.SetText(string.Join("\r\n", lines));
            });
        }

        private string GetField(string body, string name)
        {
            if (!body.Contains(name))
            {
                throw new Exception("Failed to find field in body - " + name);
            }

            return body.Split(new[] { name }, StringSplitOptions.None)[1].Split('\r')[0].Trim();
        }

        private int IndexOf(string col)
        {
            return lstResults.Columns.Cast<ColumnHeader>().First(c => c.Text == col).Index;
        }

        private bool ShowQuestion(string title, string msg)
        {
            return MessageBox.Show(this, msg + ".\nAre you sure?", title, MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes;
        }

        private async void btnGo_Click(object sender, EventArgs e)
        {
            try
            {
                btnGo.Enabled = false;
                _allSelected = true;

                SetProg(true);

                lstResults.Items.Clear();
                var folder = "Inbox";

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
                }
                else
                {
                    folder = txtFolder.Text;
                }

                SaveSettings();

                var detailsFromNote = SyncToOneNote(null, false);

                var days = Convert.ToInt32(txtDays.Text);

                Log($"Fetching mails from the last {days} days (Folder: {folder})");

                UpdateSum();

                var details = await Task.Run(() => GetItems(days, txtFolder.Text));

                foreach (var detail in details)
                {
                    var item = new ListViewItem(new string[_columns.Length])
                    {
                        Tag = detail
                    };

                    if (detailsFromNote != null)
                    {
                        var exist = detailsFromNote.FirstOrDefault(d => d.Number == detail.Number);
                        if (exist != null)
                        {
                            detail.Used = exist.Used;
                        }
                    }

                    foreach (var col in _columns)
                    {
                        string val = col.GetValue(detail) as string;
                        item.SubItems[IndexOf(col.Name)].Text = val;
                        item.Checked = true;

                        if (col.Name == "Expires")
                        {
                            var time = DateTime.ParseExact(val, "dd/M/yyyy", CultureInfo.InvariantCulture);
                            if (time < DateTime.Now || time - DateTime.Now < TimeSpan.FromDays(3))
                            {
                                item.ForeColor = Color.DarkRed;
                            }
                        }
                    }

                    lstResults.Items.Add(item);
                }

                UpdateSum();

                Log($"Found {lstResults.Items.Count} items in mail");

                if (lstResults.Items.Count > 0)
                {
                    await Task.Run(() => SyncToOneNote(details.ToArray(), true));
                }
            }
            catch (Exception ex)
            {
                Log(ex);
            }
            finally
            {
                Log("Done.");

                btnGo.Enabled = true;
                SetProg(false);
            }
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

        private List<Details> GetItems(int days, string cibusFolder)
        {
            Application outlookApplication = new Application();
            NameSpace outlookNamespace = outlookApplication.GetNamespace("MAPI");
            MAPIFolder inboxFolder = outlookNamespace.GetDefaultFolder(OlDefaultFolders.olFolderInbox);

            var items = new List<MailItem>();

            foreach (Folder folder in inboxFolder.Folders)
            {
                try
                {
                    var name = folder.Name;
                    if (!string.IsNullOrEmpty(cibusFolder) && name != cibusFolder)
                    {
                        continue;
                    }

                    foreach (MailItem item in folder.Items)
                    {
                        try
                        {
                            if (DateTime.Now - item.ReceivedTime > TimeSpan.FromDays(days))
                            {
                                continue;
                            }

                            var subject = item.Subject;
                            if (!subject.Contains("שובר על סך"))
                            {
                                continue;
                            }

                            if (item.Sender?.Name?.Contains("Cibus") == true)
                            {
                                items.Add(item);
                            }
                        }
                        catch
                        {

                        }
                    }
                }
                catch(Exception ex)
                {
                    Log(ex);
                }
            }

            return items.Select(i => new Details
            {
                Number = GetField(i.Body, $"{TitleCode}:"),
                Amount = GetField(i.Body, $"{TitleAmount}:").Split(' ')[0],
                Expires = GetField(i.Body, $"{TitleExpires} "),
                Location = GetField(i.Body, $"{TitleLocation}:"),
                Date = GetField(i.Body, $"{TitleDate}:"),
            })
            .OrderByDescending(i => i.Date)
            .ThenByDescending(i => Convert.ToInt32(i.Amount))
            .ToList();
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
            lblSum.Text = lstResults.CheckedItems.Cast<ListViewItem>()
                .Select(i => (Details)i.Tag)
                .Where(i => string.IsNullOrEmpty(i.Used))
                .Sum(i => Convert.ToInt32(i.Amount)).ToString();
        }

        private void lstResults_ItemChecked(object sender, ItemCheckedEventArgs e)
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

                foreach (ListViewItem item in lstResults.Items)
                {
                    item.Checked = _allSelected;
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

        private Details[] ParseDetails(XNamespace ns, XElement outline)
        {
            // skip header
            var rows = outline.Descendants(ns + "Row").Skip(1);

            var details = new List<Details>();

            foreach (var row in rows)
            {
                var cells = row.Descendants(ns + "Cell")
                    .Select(c => (c.Value, IsComplete(ns, c))).ToArray();

                details.Add(new Details
                {
                    Amount = cells[0].Item1,
                    Number = cells[1].Item1,
                    Date = cells[2].Item1,
                    Location = cells[3].Item1,
                    Expires = cells[4].Item1,
                    Used = cells.Any(c => c.Item2) ? "V" : ""
                });
            }

            return details.ToArray();
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

                var onenoteApp = new Microsoft.Office.Interop.OneNote.Application();
                string notebookXml;

                onenoteApp.GetHierarchy(null, HierarchyScope.hsPages, out notebookXml);

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

                string pageId;
                bool shouldUpdate = false;

                var pageNode = section.Descendants(ns + "Page").Where(n => n.Attribute("name").Value == PageName).LastOrDefault();
                string xml;

                if (pageNode == null)
                {
                    if (!update)
                    {
                        return null;
                    }

                    CreateNewPage(PageName, onenoteApp, ns, sectionId, out pageId, out xml);
                    shouldUpdate = true;
                }
                else
                {
                    pageId = pageNode.Attribute("ID").Value;
                    onenoteApp.GetPageContent(pageId, out xml, PageInfo.piAll);
                }

                var doc = XDocument.Parse(xml);
                var outline = doc.Descendants(ns + "Outline").FirstOrDefault();

                var table = outline.Descendants(ns + "Table").First();
                var content = outline.ToString();

                if (!update)
                {
                    return ParseDetails(ns, outline);
                }

                foreach (var detail in details)
                {
                    if (content.Contains(detail.Number))
                    {
                        continue;
                    }

                    shouldUpdate = true;

                    table.Add(
                    new XElement(ns + "Row",
                        BuildCell(ns, detail.Amount),
                        BuildCell(ns, detail.Number),
                        BuildCell(ns, detail.Date),
                        BuildCell(ns, detail.Location),
                        BuildCell(ns, detail.Expires)
                        ));
                }

                if (shouldUpdate)
                {
                    onenoteApp.UpdatePageContent(doc.ToString());
                }

                return null;
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

        private void SetProg(bool busy)
        {
            if (InvokeRequired)
            {
                BeginInvoke((Action)(() => SetProg(busy)));
                return;
            }

            tsProg.Style = busy ? ProgressBarStyle.Marquee : ProgressBarStyle.Continuous;
        }


        private void CreateNewPage(string pageName,
            Microsoft.Office.Interop.OneNote.Application onenoteApp,
            XNamespace ns, string sectionId,
            out string pageId,
            out string xml)
        {
            Log($"Creating new OneNote page ({pageName})");

            onenoteApp.CreateNewPage(sectionId, out pageId, NewPageStyle.npsBlankPageWithTitle);

            XElement newPage = new XElement(ns + "Page");
            newPage.SetAttributeValue("ID", pageId);
            newPage.SetAttributeValue("name", pageName);

            newPage.Add(new XElement(ns + "Title",
                            new XElement(ns + "OE",
                                new XElement(ns + "T",
                                    new XCData(pageName)))));

            var outline = new XElement(ns + "Outline",
                        new XElement(ns + "OEChildren",
                            new XElement(ns + "OE")));


            var columns = new List<XElement>();

            for (int i = 0; i < 5; ++i)
            {
                columns.Add(new XElement(ns + "Column",
                  new XAttribute("index", $"{i}"),
                  new XAttribute("width", "120")));
            }

            var row = new XElement(ns + "Row",
                BuildCell(ns, TitleAmount),
                BuildCell(ns, TitleCode),
                BuildCell(ns, TitleDate),
                BuildCell(ns, TitleLocation),
                BuildCell(ns, TitleExpires));

            var table = new XElement(ns + "Table",
                  new XAttribute("bordersVisible", "true"),
                  new XAttribute("hasHeaderRow", "true"),
                  new XElement(ns + "Columns",
                  columns
                  ),
                  row);

            outline.Add(new XElement(ns + "OEChildren",
                  new XElement(ns + "OE",
                  table)));

            newPage.Add(outline);

            xml = newPage.ToString();
        }

        private XElement BuildCell(XNamespace ns, string data)
        {
            return new XElement(ns + "Cell",
                       new XElement(ns + "OEChildren",
                       new XElement(ns + "OE",
                       new XElement(ns + "T",
                       new XCData(data)))));
        }

        private string GetObjectId(Microsoft.Office.Interop.OneNote.Application onenoteApp, XNamespace ns, string parentId, HierarchyScope scope, string objectName)
        {
            onenoteApp.GetHierarchy(parentId, scope, out string xml);

            var doc = XDocument.Parse(xml);
            var nodeName = "";

            switch (scope)
            {
                case (HierarchyScope.hsNotebooks): nodeName = "Notebook"; break;
                case (HierarchyScope.hsPages): nodeName = "Page"; break;
                case (HierarchyScope.hsSections): nodeName = "Section"; break;
                default:
                    return null;
            }

            var node = doc.Descendants(ns + nodeName).Where(n => n.Attribute("name").Value == objectName).FirstOrDefault();

            return node.Attribute("ID").Value;
        }

        private async void tsOneNote_Click(object sender, EventArgs e)
        {
            tsOneNote.Enabled = false;
            SetProg(true);

            var items = lstResults.CheckedItems.Cast<ListViewItem>()
                   .Select(i => (Details)i.Tag)
                   .ToArray();

            await RunSafeAsync(() =>
            {
                SyncToOneNote(items, true);
            });

            tsOneNote.Enabled = true;

            SetProg(false);
        }
    }

    public class Settings
    {
        public string Notebook;
        public string CibusFolder;
    }

    public class Details
    {
        public string Date;
        public string Amount;
        public string Number;
        public string Expires;
        public string Location;
        public string Used;
    }
}
