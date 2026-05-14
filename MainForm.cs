using System.Drawing;
using System.IO;
using System.Windows.Forms;

namespace VB6ToCSharp;

internal sealed class MainForm : Form
{
    private static readonly string[] VB6Extensions = ["*.bas", "*.cls", "*.ctl", "*.pag"];

    private readonly TextBox        _inputFolderBox;
    private readonly TextBox        _outputFolderBox;
    private readonly TextBox        _namespaceBox;
    private readonly CheckedListBox _fileList;
    private readonly RichTextBox    _logBox;
    private readonly Button         _convertBtn;
    private readonly Dictionary<string, string> _logPaths = new(StringComparer.OrdinalIgnoreCase);
    private readonly int _labelColW;
    private readonly int _btnColW;
    private readonly int _settingsH;

    internal MainForm()
    {
        Text            = "VB6 to CSharp Port Agent";
        MinimumSize     = new Size(720, 580);
        Size            = new Size(920, 740);
        StartPosition   = FormStartPosition.CenterScreen;
        AutoScaleMode       = AutoScaleMode.Font;
        AutoScaleDimensions = new SizeF(7F, 15F);

        _labelColW = TextRenderer.MeasureText("Out:", Font).Width + 12;
        _btnColW   = TextRenderer.MeasureText("Browse...", Font).Width + 24;
        _settingsH = (Font.Height + 14) * 3 + 14;

        _inputFolderBox  = new TextBox();
        _outputFolderBox = new TextBox();
        _namespaceBox    = new TextBox();
        _namespaceBox.Enter += (_, _) => BeginInvoke(() => _namespaceBox.SelectAll());

        _fileList = new CheckedListBox
        {
            Dock           = DockStyle.Fill,
            CheckOnClick   = true,
            IntegralHeight = false,
        };

        _logBox = new RichTextBox
        {
            Dock        = DockStyle.Fill,
            ReadOnly    = true,
            BackColor   = Color.FromArgb(28, 28, 28),
            ForeColor   = Color.LightGreen,
            Font        = new Font("Consolas", 9.25f),
            ScrollBars  = RichTextBoxScrollBars.Vertical,
            BorderStyle = BorderStyle.FixedSingle,
        };

        _convertBtn = new Button
        {
            Text     = "Convert Selected Files",
            AutoSize = true,
            Padding  = new Padding(12, 4, 12, 4),
        };

        Controls.Add(BuildRoot());
        _convertBtn.Click      += OnConvert;
        Shown                  += (_, _) => _inputFolderBox.Focus();
        _fileList.MouseDoubleClick += OnFileDoubleClick;
        _inputFolderBox.Leave  += (_, _) =>
        {
            if (Directory.Exists(_inputFolderBox.Text))
                PopulateFileList(_inputFolderBox.Text);
        };
    }

    // ─── Layout ───────────────────────────────────────────────────────────

    private TableLayoutPanel BuildRoot()
    {
        var root = new TableLayoutPanel
        {
            Dock        = DockStyle.Fill,
            RowCount    = 3,
            ColumnCount = 1,
            Padding     = new Padding(10),
        };
        root.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
        root.RowStyles.Add(new RowStyle(SizeType.Absolute, _settingsH)); // settings
        root.RowStyles.Add(new RowStyle(SizeType.Percent, 38));     // file list
        root.RowStyles.Add(new RowStyle(SizeType.Percent, 62));     // log

        root.Controls.Add(BuildSettingsPanel(), 0, 0);
        root.Controls.Add(BuildFilesPanel(),    0, 1);
        root.Controls.Add(BuildLogPanel(),      0, 2);
        return root;
    }

    private Panel BuildSettingsPanel()
    {
        var tbl = new TableLayoutPanel
        {
            Dock        = DockStyle.Fill,
            ColumnCount = 3,
            RowCount    = 3,
            Padding     = new Padding(0, 6, 0, 2),
        };
        tbl.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, _labelColW));
        tbl.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
        tbl.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, _btnColW));
        tbl.RowStyles.Add(new RowStyle(SizeType.Percent, 33.33f));
        tbl.RowStyles.Add(new RowStyle(SizeType.Percent, 33.33f));
        tbl.RowStyles.Add(new RowStyle(SizeType.Percent, 33.34f));

        var margin = new Padding(0, 2, 0, 2);
        _inputFolderBox.Dock    = DockStyle.Fill;
        _inputFolderBox.Margin  = margin;
        _outputFolderBox.Dock   = DockStyle.Fill;
        _outputFolderBox.Margin = margin;
        _namespaceBox.Dock      = DockStyle.Fill;
        _namespaceBox.Margin    = margin;

        Button MakeBtn(Action action)
        {
            var b = new Button { Text = "Browse...", Dock = DockStyle.Fill, Margin = margin };
            b.Click += (_, _) => action();
            return b;
        }

        tbl.Controls.Add(MakeLabel("In:"),  0, 0);
        tbl.Controls.Add(_inputFolderBox,   1, 0);
        tbl.Controls.Add(MakeBtn(BrowseInputFolder),  2, 0);
        tbl.Controls.Add(MakeLabel("Out:"), 0, 1);
        tbl.Controls.Add(_outputFolderBox,  1, 1);
        tbl.Controls.Add(MakeBtn(BrowseOutputFolder), 2, 1);
        tbl.Controls.Add(MakeLabel("NS:"),  0, 2);
        tbl.Controls.Add(_namespaceBox,     1, 2);
        return tbl;
    }

    private static Label MakeLabel(string text) => new()
    {
        Text      = text,
        TextAlign = ContentAlignment.MiddleRight,
        AutoSize  = false,
        Dock      = DockStyle.Fill,
        Margin    = new Padding(0, 0, 6, 0),
    };

    private Panel BuildFilesPanel()
    {
        var panel = new Panel { Dock = DockStyle.Fill, Padding = new Padding(0, 10, 0, 0) };

        var header = new TableLayoutPanel
        {
            Dock        = DockStyle.Top,
            AutoSize    = true,
            AutoSizeMode = AutoSizeMode.GrowAndShrink,
            ColumnCount = 2,
            RowCount    = 1,
        };
        header.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
        header.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));

        header.Controls.Add(new Label
        {
            Text      = "VB6 Files to Convert:",
            Dock      = DockStyle.Fill,
            TextAlign = ContentAlignment.MiddleLeft,
            Padding   = new Padding(0, 0, 0, 0),
        }, 0, 0);

        var btnRow = new FlowLayoutPanel
        {
            Dock          = DockStyle.Fill,
            FlowDirection = FlowDirection.LeftToRight,
            WrapContents  = false,
            AutoSize      = true,
            Padding       = new Padding(0, 0, 0, 0),
        };
        var selectAll   = new Button { Text = "Select All",   AutoSize = true, Margin = new Padding(0, 0, 4, 0) };
        var deselectAll = new Button { Text = "Deselect All", AutoSize = true, Margin = new Padding(0, 0, 0, 0) };
        selectAll.Click   += (_, _) => SetAllChecked(true);
        deselectAll.Click += (_, _) => SetAllChecked(false);
        btnRow.Controls.AddRange([selectAll, deselectAll]);
        header.Controls.Add(btnRow, 1, 0);

        panel.Controls.Add(_fileList);
        panel.Controls.Add(header);
        return panel;
    }

    private Panel BuildLogPanel()
    {
        var panel = new Panel { Dock = DockStyle.Fill };

        var top = new FlowLayoutPanel
        {
            Dock          = DockStyle.Top,
            AutoSize      = true,
            AutoSizeMode  = AutoSizeMode.GrowAndShrink,
            FlowDirection = FlowDirection.RightToLeft,
            WrapContents  = false,
            Padding       = new Padding(0, 4, 0, 4),
        };
        top.Controls.Add(_convertBtn);

        var logLabel = new Label
        {
            Text      = "Conversion Log:",
            Dock      = DockStyle.Top,
            Height    = Font.Height + 12,
            TextAlign = ContentAlignment.BottomLeft,
        };

        panel.Controls.Add(_logBox);
        panel.Controls.Add(logLabel);
        panel.Controls.Add(top);
        return panel;
    }

    // ─── Folder browsing ──────────────────────────────────────────────────

    private void BrowseInputFolder()
    {
        using var dlg = new FolderBrowserDialog
        {
            Description            = "Select folder containing VB6 source files",
            UseDescriptionForTitle = true,
        };
        if (Directory.Exists(_inputFolderBox.Text))
            dlg.InitialDirectory = _inputFolderBox.Text;
        if (dlg.ShowDialog(this) != DialogResult.OK) return;

        _inputFolderBox.Text = dlg.SelectedPath;
        PopulateFileList(dlg.SelectedPath);
    }

    private void BrowseOutputFolder()
    {
        using var dlg = new FolderBrowserDialog
        {
            Description        = "Select output folder for C# files",
            UseDescriptionForTitle = true,
        };
        if (Directory.Exists(_outputFolderBox.Text))
            dlg.InitialDirectory = _outputFolderBox.Text;
        if (dlg.ShowDialog(this) == DialogResult.OK)
            _outputFolderBox.Text = dlg.SelectedPath;
    }

    private void PopulateFileList(string folder, string? preSelected = null)
    {
        _fileList.Items.Clear();

        var files = VB6Extensions
            .SelectMany(ext => Directory.GetFiles(folder, ext, SearchOption.TopDirectoryOnly))
            .OrderBy(f => f)
            .ToList();

        foreach (var f in files)
        {
            bool check = preSelected is null
                || string.Equals(f, preSelected, StringComparison.OrdinalIgnoreCase);
            _fileList.Items.Add(f, isChecked: check);
        }

        _outputFolderBox.Text = Path.Combine(folder, "cs_output");
    }

    private void SetAllChecked(bool value)
    {
        for (int i = 0; i < _fileList.Items.Count; i++)
            _fileList.SetItemChecked(i, value);
    }

    // ─── Conversion ───────────────────────────────────────────────────────

    private async void OnConvert(object? sender, EventArgs e)
    {
        var selected = _fileList.CheckedItems.Cast<string>().ToList();
        if (selected.Count == 0)
        {
            MessageBox.Show("No files selected.", Text,
                MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return;
        }

        string outputDir = _outputFolderBox.Text.Trim();
        if (string.IsNullOrEmpty(outputDir))
        {
            MessageBox.Show("Please specify an output folder.", Text,
                MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return;
        }

        _convertBtn.Enabled = false;
        _logBox.Clear();
        _logPaths.Clear();
        AppendLog($"Converting {selected.Count} file(s)...\n\n", Color.Cyan);

        string baseDir = _inputFolderBox.Text.Trim();
        string ns = _namespaceBox.Text.Trim();
        int ok = 0, fail = 0, totalWarnings = 0;

        await Task.Run(() =>
        {
            try { Directory.CreateDirectory(outputDir); }
            catch (Exception ex)
            {
                AppendLog($"Cannot create output folder: {ex.Message}\n", Color.OrangeRed);
                return;
            }

            foreach (var file in selected)
            {
                string relative = (baseDir.Length > 0 && Directory.Exists(baseDir))
                    ? Path.GetRelativePath(baseDir, file)
                    : Path.GetFileName(file);

                string outFile = Path.Combine(outputDir,
                    Path.ChangeExtension(relative, ".cs"));

                try { Directory.CreateDirectory(Path.GetDirectoryName(outFile)!); }
                catch { /* ignore */ }

                bool success = ConversionEngine.ConvertFile(
                    file, outFile, ns, out int warns, msg =>
                    {
                        Color c = msg.Contains("FAILED")  ? Color.OrangeRed
                                : msg.Contains("warning") ? Color.Yellow
                                : Color.LightGreen;
                        AppendLog(msg + "\n", c);
                    });

                if (success)
                {
                    ok++;
                    totalWarnings += warns;
                    if (warns > 0)
                    {
                        string lp = Path.ChangeExtension(outFile, ".log");
                        _logPaths[file] = lp;
                    }
                }
                else fail++;
            }
        });

        string summary = $"\nDone — {ok} succeeded, {fail} failed";
        if (totalWarnings > 0) summary += $", {totalWarnings} warning(s)";
        AppendLog(summary + ".\n", fail == 0 ? Color.LightGreen : Color.OrangeRed);

        _convertBtn.Enabled = true;
    }

    private void OnFileDoubleClick(object? sender, MouseEventArgs e)
    {
        int index = _fileList.IndexFromPoint(e.Location);
        if (index < 0) return;
        string inputFile = (string)_fileList.Items[index];

        if (!_logPaths.TryGetValue(inputFile, out string? logPath) || !File.Exists(logPath)) return;

        var form = new Form
        {
            Text          = $"Log — {Path.GetFileName(logPath)}",
            Size          = new Size(700, 500),
            MinimumSize   = new Size(400, 300),
            StartPosition = FormStartPosition.CenterParent,
        };
        form.Controls.Add(new RichTextBox
        {
            Dock       = DockStyle.Fill,
            ReadOnly   = true,
            Text       = File.ReadAllText(logPath),
            Font       = new Font("Consolas", 9.25f),
            BackColor  = Color.FromArgb(28, 28, 28),
            ForeColor  = Color.Yellow,
            ScrollBars = RichTextBoxScrollBars.Vertical,
            BorderStyle = BorderStyle.None,
        });
        form.ShowDialog(this);
    }

    private void AppendLog(string text, Color color)
    {
        if (InvokeRequired) { Invoke(() => AppendLog(text, color)); return; }
        _logBox.SelectionStart  = _logBox.TextLength;
        _logBox.SelectionLength = 0;
        _logBox.SelectionColor  = color;
        _logBox.AppendText(text);
        _logBox.ScrollToCaret();
    }
}
