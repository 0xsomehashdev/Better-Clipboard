using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel; // For CancelEventArgs
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Interop;
using System.Windows.Media.Imaging;
using System.Windows.Forms; // For NotifyIcon
using System.Windows.Media; // For BitmapSource, etc.
using System.Threading; // For Thread.Sleep

namespace ClipboardApp
{
    public partial class MainWindow : Window
    {
        private const int WH_KEYBOARD_LL = 13;
        private const int WM_KEYDOWN = 0x0100;
        private const byte VK_CONTROL = 0x11;
        private const byte VK_MENU = 0x12; // Alt
        private const byte VK_C = 0x43; // C (хоткей Ctrl+Alt+C)
        private const byte VK_V = 0x56;
        private const uint KEYEVENTF_KEYUP = 0x0002;

        [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        private static extern IntPtr SetWindowsHookEx(int idHook, LowLevelKeyboardProc lpfn, IntPtr hMod, uint dwThreadId);

        [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool UnhookWindowsHookEx(IntPtr hhk);

        [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        private static extern IntPtr CallNextHookEx(IntPtr hhk, int nCode, IntPtr wParam, IntPtr lParam);

        [DllImport("kernel32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        private static extern IntPtr GetModuleHandle(string lpModuleName);

        [DllImport("user32.dll")]
        private static extern short GetAsyncKeyState(int vKey);

        [DllImport("user32.dll")]
        private static extern bool AddClipboardFormatListener(IntPtr hwnd);

        [DllImport("user32.dll")]
        private static extern bool RemoveClipboardFormatListener(IntPtr hwnd);

        [DllImport("user32.dll")]
        private static extern void keybd_event(byte bVk, byte bScan, uint dwFlags, uint dwExtraInfo);

        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool SetForegroundWindow(IntPtr hWnd);

        private delegate IntPtr LowLevelKeyboardProc(int nCode, IntPtr wParam, IntPtr lParam);

        public ObservableCollection<object> History { get; } = new ObservableCollection<object>();

        private NotifyIcon _trayIcon;
        private IntPtr _handle;
        private HwndSource _source;
        private readonly string _logPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "log.txt");
        private static IntPtr _hookID = IntPtr.Zero;
        private static LowLevelKeyboardProc _proc = HookCallback;

        public MainWindow()
        {
            InitializeComponent();
            DataContext = this;

            try
            {
                // Setup tray icon with context menu
                _trayIcon = new NotifyIcon
                {
                    Icon = LoadEmbeddedIcon(),
                    Text = "Enhanced Clipboard",
                    Visible = true
                };
                _trayIcon.Click += (sender, args) => ShowWindow();

                // Context menu for Exit
                var contextMenu = new ContextMenuStrip();
                contextMenu.Items.Add("Exit", null, (s, e) => System.Windows.Application.Current.Shutdown());
                _trayIcon.ContextMenuStrip = contextMenu;

                // Start minimized to tray
                this.WindowState = WindowState.Minimized;
                this.ShowInTaskbar = false;
                this.Hide();

                // Events
                this.Deactivated += (sender, args) => this.Hide();
                this.Closing += MainWindow_Closing;

                Log("Application started.");
            }
            catch (Exception ex)
            {
                Log($"Constructor error: {ex}");
                System.Windows.MessageBox.Show($"Startup error: {ex.Message}");
            }
        }

        private System.Drawing.Icon LoadEmbeddedIcon()
        {
            try
            {
                var assembly = Assembly.GetExecutingAssembly();
                using (var stream = assembly.GetManifestResourceStream("ClipboardApp.betterclipboarlogo.ico"))
                {
                    if (stream != null)
                    {
                        return new System.Drawing.Icon(stream);
                    }
                }
            }
            catch (Exception ex)
            {
                Log($"Error loading icon: {ex.Message}");
            }
            return System.Drawing.Icon.ExtractAssociatedIcon(Assembly.GetExecutingAssembly().Location);
        }

        protected override void OnSourceInitialized(EventArgs e)
        {
            base.OnSourceInitialized(e);
            _handle = new WindowInteropHelper(this).Handle;
            _source = PresentationSource.FromVisual(this) as HwndSource;
            _source.AddHook(WndProc); // Для клипборда

            // Set keyboard hook
            using (var curProcess = System.Diagnostics.Process.GetCurrentProcess())
            using (var curModule = curProcess.MainModule)
            {
                _hookID = SetWindowsHookEx(WH_KEYBOARD_LL, _proc, GetModuleHandle(curModule.ModuleName), 0);
                if (_hookID == IntPtr.Zero)
                {
                    Log("Keyboard hook registration failed. Error: " + Marshal.GetLastWin32Error());
                    System.Windows.MessageBox.Show("Не удалось установить глобальный хук клавиатуры. Запустите от имени администратора.");
                }
                else
                {
                    Log("Keyboard hook registered successfully.");
                }
            }

            // Start monitoring clipboard
            AddClipboardFormatListener(_handle);
        }

        protected override void OnClosed(EventArgs e)
        {
            _source.RemoveHook(WndProc);
            UnhookWindowsHookEx(_hookID);
            RemoveClipboardFormatListener(_handle);
            _trayIcon.Visible = false;
            _trayIcon.Dispose();
            base.OnClosed(e);
        }

        private IntPtr WndProc(IntPtr hwnd, int msg, IntPtr wParam, IntPtr lParam, ref bool handled)
        {
            try
            {
                if (msg == 0x031D) // WM_CLIPBOARDUPDATE
                {
                    HandleClipboardChanged();
                    handled = true;
                }
            }
            catch (Exception ex)
            {
                Log($"WndProc error: {ex}");
            }
            return IntPtr.Zero;
        }

        private static IntPtr HookCallback(int nCode, IntPtr wParam, IntPtr lParam)
        {
            try
            {
                if (nCode >= 0 && wParam == (IntPtr)WM_KEYDOWN)
                {
                    int vkCode = Marshal.ReadInt32(lParam);
                    if (vkCode == VK_C && (GetAsyncKeyState(VK_CONTROL) < 0) && (GetAsyncKeyState(VK_MENU) < 0))
                    {
                        // Ctrl+Alt+C pressed
                        var window = System.Windows.Application.Current.MainWindow as MainWindow;
                        if (window != null)
                        {
                            window.Dispatcher.Invoke(() =>
                            {
                                window.Log("Hotkey Ctrl+Alt+C triggered via hook!");
                                window.ShowWindow();
                            });
                        }
                        return (IntPtr)1; // Suppress further processing if needed (optional)
                    }
                }
            }
            catch (Exception ex)
            {
                var window = System.Windows.Application.Current.MainWindow as MainWindow;
                if (window != null)
                {
                    window.Log($"HookCallback error: {ex}");
                }
            }
            return CallNextHookEx(_hookID, nCode, wParam, lParam);
        }

        private void ShowWindow()
        {
            try
            {
                Log("ShowWindow called!");
                if (!this.IsVisible)
                {
                    this.Show();
                }
                if (this.WindowState == WindowState.Minimized)
                {
                    this.WindowState = WindowState.Normal;
                }
                var workArea = SystemParameters.WorkArea;
                this.Left = workArea.Right - this.Width;
                this.Top = workArea.Bottom - this.Height;
                this.Topmost = true;
                SetForegroundWindow(_handle); // Принудительный фокус на окно
                this.Activate();
                this.Focus();
                ClipboardListView.Focus();
                if (History.Count > 0)
                {
                    ClipboardListView.SelectedIndex = 0; // Выбор первого элемента для немедленной навигации стрелками
                }
                Log("Window focused and first item selected.");
            }
            catch (Exception ex)
            {
                Log($"ShowWindow error: {ex}");
            }
        }

        private void HandleClipboardChanged()
        {
            try
            {
                if (System.Windows.Clipboard.ContainsText())
                {
                    string text = System.Windows.Clipboard.GetText();
                    if (!string.IsNullOrEmpty(text) && (History.Count == 0 || !(History[0] is string lastText && lastText == text)))
                    {
                        History.Insert(0, text);
                    }
                }
                else if (System.Windows.Clipboard.ContainsImage())
                {
                    var bitmapSource = System.Windows.Clipboard.GetImage();
                    if (bitmapSource != null)
                    {
                        BitmapImage bitmapImage = ConvertBitmapSourceToBitmapImage(bitmapSource);
                        if (bitmapImage != null)
                        {
                            if (History.Count == 0 || !(History[0] is BitmapImage))
                            {
                                History.Insert(0, bitmapImage);
                            }
                        }
                    }
                }

                while (History.Count > 30)
                {
                    History.RemoveAt(History.Count - 1);
                }
            }
            catch (Exception ex)
            {
                Log($"HandleClipboardChanged error: {ex}");
            }
        }

        private BitmapImage ConvertBitmapSourceToBitmapImage(BitmapSource bitmapSource)
        {
            try
            {
                using (var stream = new MemoryStream())
                {
                    var encoder = new PngBitmapEncoder();
                    encoder.Frames.Add(BitmapFrame.Create(bitmapSource));
                    encoder.Save(stream);
                    stream.Position = 0;

                    var bitmapImage = new BitmapImage();
                    bitmapImage.BeginInit();
                    bitmapImage.StreamSource = stream;
                    bitmapImage.CacheOption = BitmapCacheOption.OnLoad;
                    bitmapImage.EndInit();
                    bitmapImage.Freeze();
                    return bitmapImage;
                }
            }
            catch (Exception ex)
            {
                Log($"Error converting image: {ex.Message}");
                return null;
            }
        }

        private void PasteSelected()
        {
            try
            {
                var selectedItems = ClipboardListView.SelectedItems;
                if (selectedItems.Count == 0) 
                {
                    Log("No items selected for paste.");
                    return;
                }

                List<string> texts = new List<string>();
                List<BitmapImage> images = new List<BitmapImage>();

                foreach (var item in selectedItems)
                {
                    if (item is string text)
                    {
                        texts.Add(text);
                    }
                    else if (item is BitmapImage image)
                    {
                        images.Add(image);
                    }
                }

                if (images.Count > 1 || (images.Count == 1 && texts.Count > 0))
                {
                    System.Windows.MessageBox.Show("Cannot paste multiple images or mixed text/images.", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                    Log("Paste error: invalid selection.");
                    return;
                }

                if (images.Count == 1)
                {
                    System.Windows.Clipboard.SetImage(images[0]);
                    Log("Pasted image.");
                }
                else if (texts.Count > 0)
                {
                    string joinedText = string.Join(" ", texts);
                    System.Windows.Clipboard.SetText(joinedText);
                    Log("Pasted text: " + joinedText);
                }
            }
            catch (Exception ex)
            {
                Log($"PasteSelected error: {ex}");
            }
        }

        private void ClipboardListView_PreviewKeyDown(object sender, System.Windows.Input.KeyEventArgs e)
        {
            try
            {
                if (e.Key == System.Windows.Input.Key.Escape)
                {
                    this.Hide();
                    e.Handled = true;
                }
                else if (e.Key == System.Windows.Input.Key.Enter)
                {
                    PasteSelected();
                    this.Hide();
                    Thread.Sleep(100); // Короткая пауза, чтобы фокус вернулся к предыдущему окну
                    // Simulate Ctrl+V to auto-paste
                    keybd_event(VK_CONTROL, 0, 0, 0); // Ctrl down
                    keybd_event(VK_V, 0, 0, 0); // V down
                    keybd_event(VK_V, 0, KEYEVENTF_KEYUP, 0); // V up
                    keybd_event(VK_CONTROL, 0, KEYEVENTF_KEYUP, 0); // Ctrl up
                    Log("Simulated Ctrl+V after hide.");
                    e.Handled = true;
                }
            }
            catch (Exception ex)
            {
                Log($"PreviewKeyDown error: {ex}");
            }
        }

        private void MainWindow_Closing(object sender, CancelEventArgs e)
        {
            e.Cancel = true;
            this.Hide();
        }

        private void Log(string message)
        {
            try
            {
                File.AppendAllText(_logPath, $"{DateTime.Now}: {message}\n");
            }
            catch { /* Silent fail */ }
        }
    }
}