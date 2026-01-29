using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;                 // SystemIcons
using System.IO.Ports;
using System.Linq;
using System.Management;              // WMI
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Threading;

using Forms = System.Windows.Forms;

namespace ComPortWatcher
{
    public partial class MainWindow : Window
    {
        private Dictionary<string, PortInfo> _prev = new(StringComparer.OrdinalIgnoreCase);
        private readonly DispatcherTimer _timer = new();

        private bool _reallyExit = false;
        private Forms.NotifyIcon? _tray;

        private Dictionary<string, DateTime> _insertedTime = new();

        // タイマー更新の多重実行防止
        private bool _refreshBusy = false;

        public MainWindow()
        {
            try
            {
                InitializeComponent();

                this.Closing += MainWindow_Closing;
                InitTray();

                // 起動後処理（UIが出てから）
                this.Loaded += async (_, __) =>
                {
                    try
                    {
                        // 起動時はトレイ常駐のみ
                        this.Hide();

                        // まずは SerialPort だけで即表示（軽くて確実）
                        ShowPorts(GetPortsBySerialOnly(note: "SerialPortのみ"), showDiff: false, allowNotify: false);

                        // その後WMIをバックグラウンドで取得して反映（固まってもUIは止まらない）
                        await RefreshAsync(showDiff: false, allowNotify: false);

                        // 監視開始（1秒）
                        _timer.Interval = TimeSpan.FromSeconds(1);
                        _timer.Tick += async (_, __2) => await RefreshAsync(showDiff: true, allowNotify: true);
                        _timer.Start();
                    }
                    catch (Exception loadEx)
                    {
                        LogError($"Loaded Event Error: {loadEx}");
                    }
                };
            }
            catch (Exception ex)
            {
                // 初期化エラーをログファイルに出力
                try
                {
                    var logPath = System.IO.Path.Combine(
                        System.IO.Path.GetTempPath(),
                        "ComPortWatcher_error.log"
                    );
                    System.IO.File.AppendAllText(logPath, $"{DateTime.Now:yyyy-MM-dd HH:mm:ss} Error: {ex}\n");
                }
                catch { }
                throw; // エラーを再スロー
            }
        }

        // ===== トレイ =====
        private void InitTray()
        {
            try
            {
                Icon? icon = null;
                
                try
                {
                    // アセンブリのリソースからアイコンを取得
                    var assembly = System.Reflection.Assembly.GetExecutingAssembly();
                    var resourceNames = assembly.GetManifestResourceNames();
                    
                    // icon.ico というリソースを探す
                    var iconResourceName = resourceNames.FirstOrDefault(r => 
                        r.EndsWith("icon.ico", StringComparison.OrdinalIgnoreCase));
                    
                    if (!string.IsNullOrEmpty(iconResourceName))
                    {
                        using var stream = assembly.GetManifestResourceStream(iconResourceName);
                        if (stream != null)
                        {
                            icon = new Icon(stream);
                        }
                    }
                }
                catch
                {
                    // リソース読み込み失敗時は次のフォールバックへ
                }

                // リソースがダメなら実ファイルから読み込む
                if (icon == null)
                {
                    try
                    {
                        var exeDir = System.AppContext.BaseDirectory;
                        var iconPath = System.IO.Path.Combine(exeDir, "icon.ico");
                        
                        if (System.IO.File.Exists(iconPath))
                        {
                            icon = new Icon(iconPath);
                        }
                    }
                    catch
                    {
                        // ファイルも見つからない場合はnullのまま
                    }
                }

                _tray = new Forms.NotifyIcon
                {
                    Icon = icon,
                    Text = "COM Watcher",
                    Visible = true
                };

                // 左クリックで表示
                _tray.MouseClick += (_, e) =>
                {
                    if (e.Button == Forms.MouseButtons.Left) ShowAndActivate();
                };

                // 右クリックメニュー
                var menu = new Forms.ContextMenuStrip();
                menu.Items.Add("表示", null, (_, __) => ShowAndActivate());
                menu.Items.Add(new Forms.ToolStripSeparator());
                menu.Items.Add("終了", null, (_, __) => ExitApp());
                _tray.ContextMenuStrip = menu;
            }
            catch (Exception ex)
            {
                LogError($"InitTray Error: {ex}");
            }
        }

        private static void LogError(string message)
        {
            try
            {
                var logPath = System.IO.Path.Combine(System.IO.Path.GetTempPath(), "ComPortWatcher_error.log");
                System.IO.File.AppendAllText(logPath, $"{DateTime.Now:yyyy-MM-dd HH:mm:ss} {message}\n");
            }
            catch { }
        }

        private void MainWindow_Closing(object? sender, CancelEventArgs e)
        {
            if (_reallyExit) return;

            // ×ボタンは「隠す」
            e.Cancel = true;
            this.Hide();
        }

        private void ExitApp()
        {
            _reallyExit = true;

            if (_tray != null)
            {
                _tray.Visible = false;
                _tray.Dispose();
                _tray = null;
            }

            System.Windows.Application.Current.Shutdown();
        }

        private void ShowAndActivate()
        {
            this.Show();
            this.WindowState = WindowState.Normal;
            this.Activate();
            this.Topmost = true;   // 前面に出す小技
            this.Topmost = false;
        }

        // ===== UI（更新ボタン）=====
        private async void BtnRefresh_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                await RefreshAsync(showDiff: true, allowNotify: true);
            }
            catch (Exception ex)
            {
                System.Windows.MessageBox.Show(ex.ToString(), "更新エラー", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        // ===== 更新の本体（非同期・多重実行防止）=====
        private async Task RefreshAsync(bool showDiff, bool allowNotify)
        {
            if (_refreshBusy) return;
            _refreshBusy = true;

            try
            {
                Dictionary<string, PortInfo> current;

                try
                {
                    // WMIはバックグラウンドへ（固まってもUIが止まらない）
                    current = await Task.Run(() => GetPortsSafe());
                    if (current.Count == 0)
                        current = GetPortsBySerialOnly(note: "WMIで0件");
                }
                catch
                {
                    current = GetPortsBySerialOnly(note: "WMI例外");
                }

                ShowPorts(current, showDiff, allowNotify);
            }
            finally
            {
                _refreshBusy = false;
            }
        }

private void ShowPorts(Dictionary<string, PortInfo> current, bool showDiff, bool allowNotify)
{
    // USBだけにする（差分・通知もUSBだけ）
    current = current
        .Where(kv => IsUsbSerial(kv.Value))
        .ToDictionary(kv => kv.Key, kv => kv.Value, StringComparer.OrdinalIgnoreCase);

    // 差分（ここを先に！）
    string diffText = "";
    string[] added = Array.Empty<string>();
    string[] removed = Array.Empty<string>();

    if (showDiff)
    {
        added = current.Keys.Except(_prev.Keys).OrderBy(x => x).ToArray();
        removed = _prev.Keys.Except(current.Keys).OrderBy(x => x).ToArray();

        foreach (var p in added)
            _insertedTime[p] = DateTime.Now;

        foreach (var p in removed)
            _insertedTime.Remove(p);

        if (added.Length > 0) diffText += " 追加: " + string.Join(", ", added);
        if (removed.Length > 0) diffText += " 削除: " + string.Join(", ", removed);
        if (diffText == "") diffText = " 変化なし";
    }

    // 表示対象（USBのみ）
    var usbOnly = current.Values.ToArray();

    // 初回起動時：既存ポートにも時刻を入れておく（並びが安定）
    foreach (var p in usbOnly)
        if (!_insertedTime.ContainsKey(p.PortName))
            _insertedTime[p.PortName] = DateTime.MinValue;

    // 「最後に挿されたCOMが一番上」
    var items = usbOnly
        .OrderByDescending(p => _insertedTime.TryGetValue(p.PortName, out var t) ? t : DateTime.MinValue)
        .Select(p =>
        {
            var vidpid = ExtractVidPid(p.PnpDeviceId);
            var tail = string.IsNullOrEmpty(vidpid) ? "" : $"  [{vidpid}]";
            return $"{p.PortName,-6}  {p.FriendlyName}{tail}";
        })
        .ToArray();

    ListPorts.ItemsSource = items;
    TxtStatus.Text = $"検出: {current.Count} 件 / {DateTime.Now:HH:mm:ss}{diffText}";

    // 通知（追加/削除）
    if (allowNotify && _tray != null && showDiff)
    {
        if (added.Length > 0)
        {
            var lines = added
                .Where(p => current.ContainsKey(p))
                .Select(p => $"{p}: {current[p].FriendlyName}")
                .ToArray();

            var body = string.Join("\n", lines);
            if (body.Length > 200) body = body.Substring(0, 200) + "…";

            _tray.BalloonTipTitle = "COMポートが追加されました";
            _tray.BalloonTipText = body;
            _tray.ShowBalloonTip(800);
        }

        // if (removed.Length > 0)
        // {
        //     _tray.BalloonTipTitle = "COMポートが削除されました";
        //     _tray.BalloonTipText = string.Join(", ", removed);
        //     _tray.ShowBalloonTip(1200);
        // }
    }

    _prev = current;
}

        // ===== ポート取得（WMI優先、ダメならフォールバック）=====
        private static Dictionary<string, PortInfo> GetPortsSafe()
        {
            try
            {
                var wmi = GetPortsByWmi();
                if (wmi.Count > 0) return wmi;
                return new Dictionary<string, PortInfo>(StringComparer.OrdinalIgnoreCase);
            }
            catch
            {
                return new Dictionary<string, PortInfo>(StringComparer.OrdinalIgnoreCase);
            }
        }

        private static Dictionary<string, PortInfo> GetPortsByWmi()
        {
            var result = new Dictionary<string, PortInfo>(StringComparer.OrdinalIgnoreCase);
            var re = new Regex(@"\((COM\d+)\)", RegexOptions.IgnoreCase);

            var query = "SELECT Name, PNPDeviceID FROM Win32_PnPEntity WHERE Name LIKE '%(COM%'";
            using var searcher = new ManagementObjectSearcher(query);

            foreach (ManagementObject mo in searcher.Get())
            {
                var name = mo["Name"]?.ToString() ?? "";
                var pnp = mo["PNPDeviceID"]?.ToString() ?? "";

                var m = re.Match(name);
                if (!m.Success) continue;

                var com = m.Groups[1].Value.ToUpperInvariant();
                result[com] = new PortInfo
                {
                    PortName = com,
                    FriendlyName = name,
                    PnpDeviceId = pnp
                };
            }

            return result;
        }

        private static Dictionary<string, PortInfo> GetPortsBySerialOnly(string note)
        {
            var ports = SerialPort.GetPortNames()
                .OrderBy(p => p, StringComparer.OrdinalIgnoreCase)
                .ToArray();

            var dict = new Dictionary<string, PortInfo>(StringComparer.OrdinalIgnoreCase);
            foreach (var p in ports)
            {
                dict[p] = new PortInfo
                {
                    PortName = p,
                    FriendlyName = $"(名前取得なし) {p}  [{note}]",
                    PnpDeviceId = ""
                };
            }

            return dict;
        }

        private static string ExtractVidPid(string pnpDeviceId)
        {
            if (string.IsNullOrWhiteSpace(pnpDeviceId)) return "";

            var m = Regex.Match(pnpDeviceId, @"VID_([0-9A-F]{4}).*PID_([0-9A-F]{4})", RegexOptions.IgnoreCase);
            if (!m.Success) return "";

            return $"VID_{m.Groups[1].Value.ToUpperInvariant()} PID_{m.Groups[2].Value.ToUpperInvariant()}";
        }

        private class PortInfo
        {
            public string PortName { get; set; } = "";
            public string FriendlyName { get; set; } = "";
            public string PnpDeviceId { get; set; } = "";
        }


        private static bool IsUsbSerial(PortInfo p)
        {
            // PnpDeviceIdが無い（SerialPortのみフォールバック）ときは表示する
            if (string.IsNullOrWhiteSpace(p.PnpDeviceId)) return true;

            return p.PnpDeviceId.StartsWith("USB\\", StringComparison.OrdinalIgnoreCase)
                || p.PnpDeviceId.StartsWith("USBSTOR\\", StringComparison.OrdinalIgnoreCase)
                || p.PnpDeviceId.Contains("\\VID_", StringComparison.OrdinalIgnoreCase);
        }
    }

    }
