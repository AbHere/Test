using System;
using System.Collections.Generic;
using System.IO;
using System.Text;

public sealed class CsvReader : IDisposable
{
    private readonly TextReader _reader;
    private bool _disposed;

    public CsvReader(Stream stream, Encoding? encoding = null)
    {
        _reader = new StreamReader(stream, encoding ?? Encoding.UTF8, detectEncodingFromByteOrderMarks: true);
    }

    public CsvReader(string path, Encoding? encoding = null)
    {
        _reader = new StreamReader(path, encoding ?? Encoding.UTF8, detectEncodingFromByteOrderMarks: true);
    }

    public IEnumerable<string[]> ReadRows()
    {
        var row = new List<string>();
        var field = new StringBuilder();
        bool inQuotes = false;
        int ch;

        while ((ch = _reader.Read()) != -1)
        {
            char c = (char)ch;

            if (inQuotes)
            {
                if (c == '"')
                {
                    int peek = _reader.Peek();
                    if (peek == '"') { _reader.Read(); field.Append('"'); } // escaped quote
                    else inQuotes = false;
                }
                else
                {
                    field.Append(c);
                }
            }
            else
            {
                if (c == ',')
                {
                    row.Add(field.ToString()); field.Clear();
                }
                else if (c == '\r')
                {
                    // normalize Windows newlines
                    if (_reader.Peek() == '\n') _reader.Read();
                    row.Add(field.ToString()); field.Clear();
                    yield return row.ToArray();
                    row.Clear();
                }
                else if (c == '\n')
                {
                    row.Add(field.ToString()); field.Clear();
                    yield return row.ToArray();
                    row.Clear();
                }
                else if (c == '"')
                {
                    inQuotes = true;
                }
                else
                {
                    field.Append(c);
                }
            }
        }

        // last line (no trailing newline)
        if (inQuotes)
            throw new InvalidDataException("CSV ended while inside a quoted field.");

        if (field.Length > 0 || row.Count > 0)
        {
            row.Add(field.ToString());
            yield return row.ToArray();
        }
    }

    public void Dispose()
    {
        if (_disposed) return;
        _reader.Dispose();
        _disposed = true;
    }
}
----
using System.ComponentModel;

public class TierRow : INotifyPropertyChanged
{
    string? _mnemonic, _comment, _name, _type, _status, _indentId;

    public string? Mnemonic { get => _mnemonic; set { _mnemonic = value; OnChanged(nameof(Mnemonic)); } }
    public string? Comment  { get => _comment;  set { _comment  = value; OnChanged(nameof(Comment));  } }
    public string? Name     { get => _name;     set { _name     = value; OnChanged(nameof(Name));     } }
    public string? Type     { get => _type;     set { _type     = value; OnChanged(nameof(Type));     } }

    // Import ignores this; UI shows it read-only; populated only after successful creation
    public string? IndentId { get => _indentId; set { _indentId = value; OnChanged(nameof(IndentId)); } }

    public string? Status   { get => _status;   set { _status   = value; OnChanged(nameof(Status));   } }

    public event PropertyChangedEventHandler? PropertyChanged;
    void OnChanged(string n) => PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(n));
}
------
public partial class TiersControl : UserControl
{
    private readonly BindingList<TierRow> _rows = new();
    private CancellationTokenSource? _cts;

    public TiersControl()
    {
        InitializeComponent();

        dataGridView1.AutoGenerateColumns = false;
        dataGridView1.DataSource = _rows;

        dataGridView1.Columns.Clear();
        dataGridView1.Columns.Add(new DataGridViewTextBoxColumn { DataPropertyName = nameof(TierRow.Mnemonic), HeaderText = "Mnemonic", Width = 140 });
        dataGridView1.Columns.Add(new DataGridViewTextBoxColumn { DataPropertyName = nameof(TierRow.Comment),  HeaderText = "Comment",  Width = 240 });
        dataGridView1.Columns.Add(new DataGridViewTextBoxColumn { DataPropertyName = nameof(TierRow.Name),     HeaderText = "Name",     Width = 160 });
        dataGridView1.Columns.Add(new DataGridViewTextBoxColumn { DataPropertyName = nameof(TierRow.Type),     HeaderText = "Type",     Width = 120 });
        dataGridView1.Columns.Add(new DataGridViewTextBoxColumn { DataPropertyName = nameof(TierRow.Status),   HeaderText = "Status",   Width = 120 });
        var indentCol = new DataGridViewTextBoxColumn { DataPropertyName = nameof(TierRow.IndentId), HeaderText = "Indent", Width = 150, ReadOnly = true };
        dataGridView1.Columns.Add(indentCol);

        btnUpload.Click += async (_, __) => await UploadAsync();
        btnSave.Click   += async (_, __) => await SaveAsync();
    }

    private async Task UploadAsync()
    {
        using var ofd = new OpenFileDialog { Filter = "CSV files (*.csv)|*.csv|All files (*.*)|*.*", Title = "Select CSV" };
        if (ofd.ShowDialog(this) != DialogResult.OK) return;

        _rows.Clear();

        using var reader = new CsvReader(ofd.FileName);
        var records = reader.ReadRows().ToList();
        if (records.Count == 0) return;

        var header = records[0].Select(h => (h ?? "").Trim()).ToArray();
        var data   = records.Skip(1);

        int iMnemonic = Find(header, "Mnemonic");
        int iComment  = Find(header, "Comment");
        int iName     = Find(header, "Name");
        int iType     = Find(header, "Type");
        // If CSV contains "Indent" we intentionally ignore it

        foreach (var r in data)
        {
            string? At(int i) => (i >= 0 && i < r.Length) ? r[i] : null;

            _rows.Add(new TierRow
            {
                Mnemonic = At(iMnemonic),
                Comment  = At(iComment),
                Name     = At(iName),
                Type     = At(iType),
                IndentId = null,      // force empty even if CSV had a value
                Status   = "Pending"
            });
        }

        static int Find(string[] hdr, string name) =>
            Array.FindIndex(hdr, h => h.Equals(name, StringComparison.OrdinalIgnoreCase));
    }

    private async Task SaveAsync()
    {
        if (_rows.Count == 0)
        {
            MessageBox.Show(this, "No rows to process.", "Info", MessageBoxButtons.OK, MessageBoxIcon.Information);
            return;
        }

        btnSave.Enabled = btnUpload.Enabled = false;
        _cts?.Dispose();
        _cts = new CancellationTokenSource();
        var token = _cts.Token;

        using var wait = new WaitingForm
        {
            TextLine = "Creating tiers…",
            CanCancel = true,
            Max = _rows.Count,
            OnCancel = () => _cts?.Cancel()
        };
        wait.Show(this);

        try
        {
            int completed = 0;

            await Task.Run(async () =>
            {
                foreach (var row in _rows)
                {
                    if (token.IsCancellationRequested) break;

                    UpdateRow(row, status: "Working…", indentId: null);
                    try
                    {
                        var indent = await CreateTierAsync(row, token); // <-- your real call
                        UpdateRow(row, status: "Created", indentId: indent);
                    }
                    catch (OperationCanceledException)
                    {
                        UpdateRow(row, status: "Canceled", indentId: null);
                        break;
                    }
                    catch (Exception ex)
                    {
                        UpdateRow(row, status: "Failed: " + Short(ex.Message), indentId: null);
                    }

                    completed++;
                    wait.SafeStep(completed);
                }
            }, token);
        }
        finally
        {
            wait.Close();
            btnSave.Enabled = btnUpload.Enabled = true;
        }

        // UI-safe row update
        void UpdateRow(TierRow r, string? status, string? indentId)
        {
            if (InvokeRequired) { BeginInvoke(new Action(() => UpdateRow(r, status, indentId))); return; }
            if (status  != null) r.Status  = status;
            if (indentId!= null) r.IndentId= indentId;   // stays null if we pass null
        }

        static string Short(string s) => s.Length > 120 ? s[..117] + "…" : s;
    }

    // Simulate your creation logic — replace with repository/service; return generated indent
    private static async Task<string> CreateTierAsync(TierRow row, CancellationToken ct)
    {
        // Validate required fields
        if (string.IsNullOrWhiteSpace(row.Mnemonic))
            throw new InvalidOperationException("Mnemonic is required.");

        // TODO: call your domain service to create the 'tier/chair'
        await Task.Delay(400, ct); // simulate I/O

        return "IND-" + Guid.NewGuid().ToString("N")[..8].ToUpperInvariant();
    }
}
------
using System;
using System.Windows.Forms;

public class WaitingForm : Form
{
    private readonly ProgressBar _bar = new() { Dock = DockStyle.Top, Style = ProgressBarStyle.Continuous, Height = 24 };
    private readonly Label _label = new() { Dock = DockStyle.Top, AutoSize = false, Height = 28, TextAlign = System.Drawing.ContentAlignment.MiddleLeft };
    private readonly Button _cancel = new() { Dock = DockStyle.Top, Height = 28, Text = "Cancel" };

    public Action? OnCancel { get; set; }
    public bool CanCancel { get => _cancel.Enabled; set => _cancel.Enabled = value; }

    public int Max { get => _bar.Maximum; set { _bar.Maximum = value; _bar.Value = 0; } }
    public string TextLine { get => _label.Text; set => _label.Text = value; }

    public WaitingForm()
    {
        FormBorderStyle = FormBorderStyle.FixedDialog;
        StartPosition = FormStartPosition.CenterParent;
        ControlBox = false;
        MinimizeBox = MaximizeBox = false;
        Width = 420; Height = 140;
        Padding = new Padding(12);

        _cancel.Click += (_, __) => OnCancel?.Invoke();

        Controls.Add(_cancel);
        Controls.Add(_bar);
        Controls.Add(_label);
    }

    public void SafeStep(int value)
    {
        if (InvokeRequired) BeginInvoke(new Action<int>(SafeStep), value);
        else _bar.Value = Math.Min(value, _bar.Maximum);
    }
}
