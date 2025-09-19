using System;
using System.Threading;

internal static class StaRunner
{
    /// Run `work` on a dedicated STA thread. If `cancel` is signaled we call `onCancel`.
    /// After `hardCancelAfter`, we abort the thread (last resort on .NET Fx).
    public static T Run<T>(Func<T> work, CancellationToken cancel, Action onCancel, TimeSpan hardCancelAfter)
    {
        if (work == null) throw new ArgumentNullException("work");

        T result = default(T);
        Exception error = null;
        using (var done = new ManualResetEventSlim(false))
        {
            Thread worker = null;
            worker = new Thread(() =>
            {
                try { result = work(); }
                catch (Exception ex) { error = ex; }
                finally { done.Set(); }
            });

            worker.IsBackground = true;
            worker.SetApartmentState(ApartmentState.STA);
            worker.Start();

            // cooperative cancel
            using (cancel.Register(() => { try { onCancel?.Invoke(); } catch { } }))
            {
                // hard cancel timer (only if needed)
                using (var killer = new Timer(_ =>
                {
                    try
                    {
                        if (!done.IsSet && worker.IsAlive)
                            worker.Abort(); // .NET Framework only; last resort
                    }
                    catch { }
                }, null, hardCancelAfter, Timeout.InfiniteTimeSpan))
                {
                    done.Wait(); // wait worker completion/abort
                }
            }

            if (error != null) throw error;
            return result;
        }
    }
}

......
public static class SophisPartyFactory
{
    private static readonly Dictionary<ParentOption,int> ParentCodes = new Dictionary<ParentOption,int>
    {
        { ParentOption.INT,  10007538 },
        { ParentOption.EXT,  10007539 },
        { ParentOption.AFB,  10007540 },
        { ParentOption.ISDA, 10007545 },
    };

    private static readonly TimeSpan HardKillAfter = TimeSpan.FromSeconds(20); // tune

    public static Task<string> CreateAsync(TierCreateRequest req, CancellationToken ct)
    {
        if (req == null) throw new ArgumentNullException("req");
        if (string.IsNullOrWhiteSpace(req.Mnemonic)) throw new ArgumentException("Mnemonic is required.");
        if (req.Institutions == InstitutionFlags.None) throw new ArgumentException("At least one Institution Type is required.");

        return Task.Run(() =>
        {
            // locals so onCancel can reach them if your API supports cooperative cancel
            CSNThirdPartyDlg dialog = null;

            return StaRunner.Run<string>(
                work: () =>
                {
                    ct.ThrowIfCancellationRequested();

                    dialog = new CSNThirdPartyDlg();
                    dialog.SetReference(req.Mnemonic);

                    var inst = req.Institutions;
                    if ((inst & InstitutionFlags.Counterparty)   != 0) dialog.SetCounterparty(true);
                    if ((inst & InstitutionFlags.Broker)         != 0) dialog.SetBroker(true);
                    if ((inst & InstitutionFlags.ClearingHouse)  != 0) dialog.SetclearingHouse(true);
                    if ((inst & InstitutionFlags.Group)          != 0) dialog.SetGroup(true);
                    if ((inst & InstitutionFlags.ExecutionVenue) != 0) dialog.SetExecutionVenue(true);
                    if ((inst & InstitutionFlags.Depositary)     != 0) dialog.SetDepository(true);
                    if ((inst & InstitutionFlags.Customer)       != 0) dialog.SetCustomer(true);
                    if ((inst & InstitutionFlags.PSET)           != 0) dialog.SetPSET(true);
                    if ((inst & InstitutionFlags.ClearingMember) != 0) dialog.SetClearingMember(true);
                    if ((inst & InstitutionFlags.TradeRepository)!= 0) dialog.SetTradeRepositary(true);

                    var cat = req.Categories;
                    if ((cat & CategoryFlags.Corporate)   != 0) dialog.SetCorporate(true);
                    if ((cat & CategoryFlags.Exchange)    != 0) dialog.SetExchange(true);
                    if ((cat & CategoryFlags.Institution) != 0) dialog.SetInstitution(true);
                    if ((cat & CategoryFlags.Bank)        != 0) dialog.SetBank(true);
                    if ((cat & CategoryFlags.Other)       != 0) dialog.Setother(true);

                    var rep = req.Reporting;
                    if ((rep & ReportingOptions.GrossPrice)   != 0) dialog.SetGrossPrice(true);
                    if ((rep & ReportingOptions.AveragePrice) != 0) dialog.SetAveragePrice(true);
                    if ((rep & ReportingOptions.MarketTax)    != 0) dialog.SetMarketTax(true);

                    if (req.IsEntity.HasValue) dialog.SetEntity(req.IsEntity.Value);

                    if (req.Parent != ParentOption.None)
                    {
                        int code; if (ParentCodes.TryGetValue(req.Parent, out code))
                            dialog.SetParent(code); // int overload
                    }

                    return dialog.Save(); // blocking call runs on STA worker
                },
                cancel: ct,
                onCancel: () =>
                {
                    try
                    {
                        // If your dialog supports it, do cooperative cancel here:
                        // dialog?.Cancel(); or dialog?.Close();
                    }
                    catch { }
                },
                hardCancelAfter: HardKillAfter
            );
        }, ct);
    }
}

.....

using System;
using System.Threading.Tasks;
using System.Windows.Forms;

internal static class Guarded
{
    public static async Task RunAsync(Control owner, Func<Task> body)
    {
        try { await body().ConfigureAwait(true); }
        catch (OperationCanceledException)
        {
            // optional UI message – often we stay silent
        }
        catch (Exception ex)
        {
            Log.Write("[HANDLER] " + ex);
            MessageBox.Show(owner, "Operation failed:\r\n" + ex.Message,
                "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
    }
}



.....

private async void btnUpload_Click(object s, EventArgs e)
    => await Guarded.RunAsync(this, () => UploadAsync());

private async void btnSave_Click(object s, EventArgs e)
    => await Guarded.RunAsync(this, () => SaveAsync());

......

using System;
using System.Threading;

internal static class StaRunner
{
    public static T Run<T>(Func<T> work, CancellationToken cancel, Action onCancel, TimeSpan hardCancelAfter)
    {
        if (work == null) throw new ArgumentNullException("work");

        T result = default(T);
        Exception error = null;
        using (var done = new ManualResetEventSlim(false))
        {
            Thread worker = new Thread(() =>
            {
                try { result = work(); }
                catch (Exception ex) { error = ex; }
                finally { done.Set(); }
            });
            worker.IsBackground = true;
            worker.SetApartmentState(ApartmentState.STA);
            worker.Start();

            using (cancel.Register(() => { try { onCancel?.Invoke(); } catch { } }))
            using (var killer = new Timer(_ =>
            {
                try
                {
                    if (!done.IsSet && worker.IsAlive)
                        worker.Abort(); // .NET Fx only – last resort to protect host UI
                }
                catch { }
            }, null, hardCancelAfter, Timeout.InfiniteTimeSpan))
            {
                done.Wait();
            }

            if (error != null) throw error;
            return result;
        }
    }
}


......

private static readonly TimeSpan HardKillAfter = TimeSpan.FromSeconds(20);

public static Task<string> CreateAsync(TierCreateRequest req, CancellationToken ct)
{
    if (req == null) throw new ArgumentNullException("req");
    if (string.IsNullOrWhiteSpace(req.Mnemonic)) throw new ArgumentException("Mnemonic is required.");
    if (req.Institutions == InstitutionFlags.None) throw new ArgumentException("At least one Institution Type is required.");

    return Task.Run(() =>
    {
        CSNThirdPartyDlg dialog = null;

        return StaRunner.Run<string>(
            work: () =>
            {
                ct.ThrowIfCancellationRequested();

                dialog = new CSNThirdPartyDlg();
                dialog.SetReference(req.Mnemonic);

                // ... your existing setters for Institutions, Categories, Reporting, IsEntity, Parent(int) ...

                return dialog.Save(); // runs on STA worker, not host UI
            },
            cancel: ct,
            onCancel: () =>
            {
                try { /* dialog?.Cancel() or dialog?.Close() if available */ } catch { }
            },
            hardCancelAfter: HardKillAfter
        );
    }, ct);
}

....

private CancellationTokenSource _cts;

private async Task SaveAsync()
{
    if (_rows.Count == 0) { MessageBox.Show(this, "No rows to process."); return; }

    // quick validation
    foreach (var r in _rows)
        if (string.IsNullOrWhiteSpace(r.Mnemonic) || r.Institutions == InstitutionFlags.None)
            r.Status = "Failed: missing Mnemonic/InstitutionTypes";

    btnSave.Enabled = false; btnUpload.Enabled = false;

    _cts?.Dispose();
    _cts = new CancellationTokenSource();
    var token = _cts.Token;

    using (var wait = new WaitingForm())
    {
        wait.TextLine = "Creating tiers…";
        wait.Max = _rows.Count;
        wait.CanCancel = true;
        wait.OnCancel = () => _cts?.Cancel();
        wait.Show(this);

        try
        {
            var dop = Math.Max(1, Environment.ProcessorCount - 1);
            using (var gate = new System.Threading.SemaphoreSlim(dop))
            {
                int completed = 0;
                var tasks = new List<Task>();

                foreach (var row in _rows.Where(r => r.Status == "Pending"))
                {
                    await gate.WaitAsync(token).ConfigureAwait(true);

                    tasks.Add(Task.Run(async () =>
                    {
                        try
                        {
                            token.ThrowIfCancellationRequested();
                            UpdateRow(row, "Working…", null);

                            using (var rowCts = CancellationTokenSource.CreateLinkedTokenSource(token))
                            {
                                rowCts.CancelAfter(TimeSpan.FromSeconds(30)); // per-row timeout

                                var req = new TierCreateRequest
                                {
                                    Mnemonic     = row.Mnemonic,
                                    Institutions = row.Institutions,
                                    Categories   = row.Categories,
                                    Reporting    = row.Reporting,
                                    IsEntity     = row.IsEntity,
                                    Parent       = row.Parent
                                };

                                var indent = await SophisPartyFactory.CreateAsync(req, rowCts.Token).ConfigureAwait(false);
                                UpdateRow(row, "Created", indent);
                            }
                        }
                        catch (OperationCanceledException)
                        {
                            UpdateRow(row, token.IsCancellationRequested ? "Canceled" : "Timeout", null);
                        }
                        catch (Exception exRow)
                        {
                            var msg = exRow.Message ?? "Error";
                            if (msg.Length > 200) msg = msg.Substring(0, 197) + "…";
                            row.Error = exRow.ToString();
                            Log.Write("[ROW] " + exRow);
                            UpdateRow(row, "Failed: " + msg, null);
                        }
                        finally
                        {
                            System.Threading.Interlocked.Increment(ref completed);
                            wait.SafeStep(completed);
                            gate.Release();
                        }
                    }, token));
                }

                await Task.WhenAll(tasks).ConfigureAwait(true);
            }
        }
        finally
        {
            wait.Close();
            btnSave.Enabled = true; btnUpload.Enabled = true;
            _cts.Dispose(); _cts = null;
        }
    }
}


....

using System;
using System.Threading.Tasks;
using System.Windows.Forms;

internal static class SafetyNet
{
    private static bool _installed;

    public static void Install(Action<string> log, bool showMessageBoxOnUiCrash = true)
    {
        if (_installed || log == null) return;
        _installed = true;

        // WinForms UI thread exceptions (works even when you don't own Program.cs)
        Application.ThreadException += (s, e) =>
        {
            try { log("[UI] " + e.Exception); } catch { }
            if (showMessageBoxOnUiCrash)
                MessageBox.Show("An unexpected error occurred.\r\n" + e.Exception.Message,
                    "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            // Do not rethrow – swallow so host window stays alive.
        };

        // Background threads (may still terminate process if truly unhandled,
        // so our main protection is to catch at our boundaries)
        AppDomain.CurrentDomain.UnhandledException += (s, e) =>
        {
            try { log("[APPDOMAIN] " + (e.ExceptionObject as Exception)); } catch { }
        };

        // Task exceptions that weren't awaited
        TaskScheduler.UnobservedTaskException += (s, e) =>
        {
            try { log("[TASK] " + e.Exception); } catch { }
            e.SetObserved(); // avoid escalation to AppDomain crash
        };
    }
}







