using System.Globalization;
using System.Net.Sockets;
using System.Text;

namespace Adressen.cls;

public enum FritzCallType
{
    Ring, Call, Connect, Disconnect
}

public sealed class FritzCallEvent
{
    public DateTime Timestamp
    {
        get; init;
    }
    public FritzCallType Type
    {
        get; init;
    }
    public string ConnectionId { get; init; } = "";
    public string CallerNumber { get; init; } = "";  // bei RING: externe Nr.
    public string CalledNumber { get; init; } = "";  // bei RING: deine Nebenstelle
    public string Extension { get; init; } = "";  // bei CALL/CONNECT
    public int DurationSeconds
    {
        get; init;
    }        // nur bei DISCONNECT

    public bool IsAnonymous => string.IsNullOrWhiteSpace(CallerNumber) || CallerNumber.Equals("anonymous", StringComparison.OrdinalIgnoreCase);
}

public sealed class FritzCallMonitor(string host = "192.168.30.1", int port = 1012, TimeSpan? reconnectDelay = null) : IDisposable, IAsyncDisposable
{
    private readonly string _host = host;
    private readonly int _port = port;
    private readonly TimeSpan _reconnectDelay = reconnectDelay ?? TimeSpan.FromSeconds(5);
    private CancellationTokenSource? _cts;
    private Task? _monitorTask;

    public event EventHandler<FritzCallEvent>? CallEventReceived;
    public event EventHandler<bool>? ConnectionChanged;  // true = verbunden

    public void Start()
    {
        if (_cts is { IsCancellationRequested: false }) { return; }
        _cts = new CancellationTokenSource();
        _monitorTask = Task.Run(() => RunAsync(_cts.Token));
    }

    public async Task StopAsync()
    {
        if (_cts is null) { return; }
        await _cts.CancelAsync();
        if (_monitorTask is not null) { await _monitorTask.ConfigureAwait(false); }
    }

    private async Task RunAsync(CancellationToken ct)
    {
        while (!ct.IsCancellationRequested)
        {
            var wasConnected = false;
            try
            {
                using var client = new TcpClient();
                await client.ConnectAsync(_host, _port, ct);
                wasConnected = true;
                ConnectionChanged?.Invoke(this, true);

                using var reader = new StreamReader(client.GetStream(), Encoding.UTF8);
                while (!ct.IsCancellationRequested)
                {
                    var line = await reader.ReadLineAsync(ct);
                    if (line is null) { break; }
                    if (TryParse(line, out var evt)) { CallEventReceived?.Invoke(this, evt!); }
                }
            }
            catch (OperationCanceledException) { break; }
            catch { }
            finally
            {
                if (wasConnected) { ConnectionChanged?.Invoke(this, false); }
            }

            try { await Task.Delay(_reconnectDelay, ct); }
            catch (OperationCanceledException) { break; }
        }
    }

    private static bool TryParse(string line, out FritzCallEvent? result)
    {
        result = null;
        var p = line.Split(';');
        if (p.Length < 4) { return false; }

        // FritzBox verwendet zweistellige Jahreszahl: "25.05.26 10:30:00"
        if (!DateTime.TryParseExact(p[0], "dd.MM.yy HH:mm:ss", CultureInfo.InvariantCulture, DateTimeStyles.None, out var ts)) { return false; }
        result = p[1].ToUpperInvariant() switch
        {
            "RING" when p.Length >= 5 => new FritzCallEvent { Timestamp = ts, Type = FritzCallType.Ring, ConnectionId = p[2], CallerNumber = p[3], CalledNumber = p[4] },
            "CALL" when p.Length >= 6 => new FritzCallEvent { Timestamp = ts, Type = FritzCallType.Call, ConnectionId = p[2], Extension = p[3], CallerNumber = p[4], CalledNumber = p[5] },
            "CONNECT" when p.Length >= 5 => new FritzCallEvent { Timestamp = ts, Type = FritzCallType.Connect, ConnectionId = p[2], Extension = p[3], CallerNumber = p[4] },
            "DISCONNECT" when p.Length >= 4 => new FritzCallEvent { Timestamp = ts, Type = FritzCallType.Disconnect, ConnectionId = p[2], DurationSeconds = int.TryParse(p[3], out var d) ? d : 0 },
            _ => null
        };

        return result is not null;
    }

    // Synchron: nur Signal senden, kein Warten
    public void Dispose()
    {
        _cts?.Cancel();
        _cts?.Dispose();
        // _monitorTask läuft aus sobald ct.IsCancellationRequested → kein Warten nötig
    }

    // Asynchron: sauber warten bis der Task wirklich fertig ist
    public async ValueTask DisposeAsync()
    {
        await StopAsync();
        _cts?.Dispose();
    }

}