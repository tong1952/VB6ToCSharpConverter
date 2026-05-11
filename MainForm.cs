using System.Drawing;
using System.IO;
using System.Windows.Forms;

namespace VB6ToCSharp;

internal sealed class MainForm : Form
{
    private static readonly string[] VB6Extensions = ["*.bas", "*.cls", "*.frm", "*.ctl", "*.pag"];

    private readonly TextBox        _inputFolderBox;
    private readonly TextBox        _outputFolderBox;
    private readonly CheckedListBox _fileList;
    private readonly RichTextBox    _logBox;
    private readonly Button         _convertBtn;
    private readonly Dictionary<string, string> _logPaths = new(StringComparer.OrdinalIgnoreCase);

    internal MainForm()
    {
        Text          = "VB6 to CSharp Port Agent";
        MinimumSize   = new Size(720, 580);
        Size          = new Size(920, 740);
        StartPosition = FormStartPosition.CenterScreen;

        _inputFolderBox  = new TextBox();
        _outputFolderBox = new TextBox();

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
            Text   = "Convert Selected Files",
            Width  = 200,
            Height = 36,
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
        root.RowStyles.Add(new RowStyle(SizeType.AutoSize));        // settings
        root.RowStyles.Add(new RowStyle(SizeType.Percent, 38));     // file list
        root.RowStyles.Add(new RowStyle(SizeType.Percent, 62));     // log

        root.Controls.Add(BuildSettingsPanel(), 0, 0);
        root.Controls.Add(BuildFilesPanel(),    0, 1);
        root.Controls.Add(BuildLogPanel(),      0, 2);
        return root;
    }

    // Two rows: Input and Output — each row is a panel with label + textbox + button
    // at the same pixel height so nothing is off-centre relative to the text box border.
    private Panel BuildSettingsPanel()
    {
        var outer = new Panel { Dock = DockStyle.Fill, AutoSize = true,
                                Padding = new Padding(0, 20, 0, 8) };

        var rowIn  = MakeFolderRow("In:",  _inputFolderBox,  BrowseInputFolder);
        var rowOut = MakeFolderRow("Out:", _outputFolderBox, BrowseOutputFolder);

        rowIn.Dock  = DockStyle.Top;
        rowOut.Dock = DockStyle.Top;

        // DockStyle.Top: last added is topmost — add in reverse order
        outer.Controls.Add(rowOut);
        outer.Controls.Add(rowIn);
        return outer;
    }

    private static Panel MakeFolderRow(string labelText, TextBox box, Action onBrowse)
    {
        const int LabelW = 60;
        const int BtnW   = 80;
        const int Pad    = 8;   // gap between controls

        // Row height driven by the TextBox's own preferred height
        var row = new Panel { Height = 32 };

        var lbl = new Label
        {
            Text      = labelText,
            TextAlign = ContentAlignment.MiddleRight,
            Left = 0, Width = LabelW,
            // height / top set in Resize
        };

        box.Left   = LabelW + Pad;
        box.Width  = 100;   // placeholder — corrected in Resize
        box.Anchor = AnchorStyles.Left | AnchorStyles.Right | AnchorStyles.Top;

        var btn = new Button
        {
            Text   = "Browse...",
            Width  = BtnW,
            Anchor = AnchorStyles.Right | AnchorStyles.Top,
        };
        btn.Click += (_, _) => onBrowse();

        row.Controls.AddRange([lbl, box, btn]);

        row.Resize += (_, _) =>
        {
            int tbH  = box.PreferredHeight;
            int top  = (row.ClientSize.Height - tbH) / 2;

            lbl.Top    = top;
            lbl.Height = tbH;

            box.Top    = top;

            btn.Width  = BtnW;
            btn.Height = tbH;
            btn.Left   = row.ClientSize.Width - BtnW;
            btn.Top    = top;

            box.Width  = btn.Left - box.Left - Pad;
        };

        return row;
    }

    private Panel BuildFilesPanel()
    {
        var panel = new Panel { Dock = DockStyle.Fill, Padding = new Padding(0, 10, 0, 0) };

        var header = new TableLayoutPanel
        {
            Dock        = DockStyle.Top,
            Height      = 30,
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
        var selectAll   = new Button { Text = "Select All",   Width = 120, Height = 30, Margin = new Padding(0, 0, 4, 0) };
        var deselectAll = new Button { Text = "Deselect All", Width = 120, Height = 30, Margin = new Padding(0, 0, 0, 0) };
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
            Height        = 44,
            FlowDirection = FlowDirection.RightToLeft,
            WrapContents  = false,
            Padding       = new Padding(0, 0, 0, 0),
        };
        top.Controls.Add(_convertBtn);

        var logLabel = new Label
        {
            Text      = "Conversion Log:",
            Dock      = DockStyle.Top,
            Height    = 28,
            TextAlign = ContentAlignment.BottomLeft,
            Padding   = new Padding(0, 0, 0, 0),
        };

        panel.Controls.Add(_logBox);
        panel.Controls.Add(logLabel);
        panel.Controls.Add(top);
        return panel;
    }

    // ─── Folder browsing ──────────────────────────────────────────────────

    private void BrowseInputFolder()
    {
        using var dlg = new OpenFileDialog
        {
            Title            = "Select any VB6 file in the source folder",
            Filter           = "VB6 Files (*.bas;*.cls;*.frm;*.ctl;*.pag)|*.bas;*.cls;*.frm;*.ctl;*.pag|All Files (*.*)|*.*",
            Multiselect      = false,
            CheckFileExists  = false,
            InitialDirectory = Directory.Exists(_inputFolderBox.Text) ? _inputFolderBox.Text : string.Empty,
        };
        if (dlg.ShowDialog(this) != DialogResult.OK) return;

        string folder = Path.GetDirectoryName(dlg.FileName)!;
        _inputFolderBox.Text = folder;
        PopulateFileList(folder, dlg.FileName);
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
            .SelectMany(ext => Directory.GetFiles(folder, ext, SearchOption.AllDirectories))
            .OrderBy(f => f)
            .ToList();

        foreach (var f in files)
        {
            bool check = preSelected is null
                || string.Equals(f, preSelected, StringComparison.OrdinalIgnoreCase);
            _fileList.Items.Add(f, isChecked: check);
        }

        if (!Directory.Exists(_outputFolderBox.Text))
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
                    file, outFile, "Converted", out int warns, msg =>
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
