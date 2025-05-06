using System.Diagnostics;
using System.Diagnostics.Eventing.Reader;
using System.IO;
using System.IO.Ports;
using System.Security.Cryptography;
using System.Security.Policy;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using System.Windows.Threading;
using CSCore.CoreAudioAPI;
using Microsoft.Win32;
using Newtonsoft.Json.Linq;

namespace Smart_Audio_Mixer_Config
{
    public partial class MainWindow : Window
    {
        private AudioMeterInformation AudioBar1Session = null;
        private AudioMeterInformation AudioBar2Session = null;
        private AudioMeterInformation AudioBar3Session = null;
        private AudioMeterInformation AudioBar4Session = null;
        private AudioMeterInformation AudioBar5Session = null;

        private static AudioSessionControl2 cachedSessionControl1 = null;
        private static AudioSessionControl2 cachedSessionControl2 = null;
        private static AudioSessionControl2 cachedSessionControl3 = null;
        private static AudioSessionControl2 cachedSessionControl4 = null;
        private static AudioSessionControl2 cachedSessionControl5 = null;

        private DispatcherTimer _audioBarTimer;
        private DispatcherTimer _hardwareScanTimer;
        private DispatcherTimer _backgroundTaskScanTimer;

        public MainWindow()
        {
            InitializeComponent();
            AsyncLoading();
        }

        private async void AsyncLoading()
        {
            await PopulateComboBoxes();
            await Task.Delay(500); // Delay to wait for Combo Boxes to populate
            await LoadSettings();

            string AudioBar1AppString = AudioBar1AppComboBox.SelectedItem.ToString();
            string AudioBar2AppString = AudioBar2AppComboBox.SelectedItem.ToString();
            string AudioBar3AppString = AudioBar3AppComboBox.SelectedItem.ToString();
            string AudioBar4AppString = AudioBar4AppComboBox.SelectedItem.ToString();
            string AudioBar5AppString = AudioBar5AppComboBox.SelectedItem.ToString();

            GetAudioSession(AudioBar1AppString, session => AudioBar1Session = session, "1");
            GetAudioSession(AudioBar2AppString, session => AudioBar2Session = session, "2");
            GetAudioSession(AudioBar3AppString, session => AudioBar3Session = session, "3");
            GetAudioSession(AudioBar4AppString, session => AudioBar4Session = session, "4");
            GetAudioSession(AudioBar5AppString, session => AudioBar5Session = session, "5");

            StartAudioMonitorThread(AudioBar1AppString, AudioBar1VolumeRectangle, () => AudioBar1Session);
            StartAudioMonitorThread(AudioBar2AppString, AudioBar2VolumeRectangle, () => AudioBar2Session);
            StartAudioMonitorThread(AudioBar3AppString, AudioBar3VolumeRectangle, () => AudioBar3Session);
            StartAudioMonitorThread(AudioBar4AppString, AudioBar4VolumeRectangle, () => AudioBar4Session);
            StartAudioMonitorThread(AudioBar5AppString, AudioBar5VolumeRectangle, () => AudioBar5Session);

            GetAudioBarValueSession(AudioBar1AppString, 1);
            GetAudioBarValueSession(AudioBar2AppString, 2);
            GetAudioBarValueSession(AudioBar3AppString, 3);
            GetAudioBarValueSession(AudioBar4AppString, 4);
            GetAudioBarValueSession(AudioBar5AppString, 5);

            _audioBarTimer = new DispatcherTimer();
            _audioBarTimer.Interval = TimeSpan.FromMilliseconds(10);
            _audioBarTimer.Tick += AudioBarTimer_Tick;
            _audioBarTimer.Start();

            _hardwareScanTimer = new DispatcherTimer();
            _hardwareScanTimer.Interval = TimeSpan.FromMilliseconds(500);
            _hardwareScanTimer.Tick += HardwareScanTimer_Tick;
            _hardwareScanTimer.Start();

            _backgroundTaskScanTimer = new DispatcherTimer();
            _backgroundTaskScanTimer.Interval = TimeSpan.FromMilliseconds(500);
            _backgroundTaskScanTimer.Tick += BackgroundTaskScan_Tick;
            _backgroundTaskScanTimer.Start();

        }

        private async Task LoadSettings()
        {
            string filePath = System.IO.Path.Combine(Environment.CurrentDirectory, "config.json"); // Looks for config.json in the same directory as the executable
            if (File.Exists(filePath))
            {
                JObject settings = JObject.Parse(System.IO.File.ReadAllText(filePath));

                string AudioBar1AppString = (string)settings["appSource"]["AudioBar1"];
                string AudioBar2AppString = (string)settings["appSource"]["AudioBar2"];
                string AudioBar3AppString = (string)settings["appSource"]["AudioBar3"];
                string AudioBar4AppString = (string)settings["appSource"]["AudioBar4"];
                string AudioBar5AppString = (string)settings["appSource"]["AudioBar5"];

                string DeviceVIDString = (string)settings["device"]["DeviceVID"];
                string DevicePIDString = (string)settings["device"]["DevicePID"];

                await Task.Run(() =>
                {
                    Dispatcher.Invoke(() =>
                    {
                        if (!string.IsNullOrEmpty(AudioBar1AppString))
                        {
                            bool found = false;
                            foreach (var item in AudioBar1AppComboBox.Items)
                            {
                                var itemString = item.ToString().ToLower();
                                if (itemString.Contains(AudioBar1AppString.ToLower()))
                                {
                                    AudioBar1AppComboBox.SelectedItem = AudioBar1AppString.ToLower();
                                    found = true;
                                    break;
                                }
                            }

                            if (!found)
                            {
                                string notDetected = $"App Not Detected ( {AudioBar1AppString} )";
                                if (!AudioBar1AppComboBox.Items.Contains(notDetected))
                                    AudioBar1AppComboBox.Items.Insert(0, notDetected);

                                AudioBar1AppComboBox.SelectedIndex = 0;
                            }
                        }

                    });
                });
                await Task.Run(() =>
                {
                    Dispatcher.Invoke(() =>
                    {
                        if (!string.IsNullOrEmpty(AudioBar2AppString))
                        {
                            bool found = false;
                            foreach (var item in AudioBar2AppComboBox.Items)
                            {
                                var itemString = item.ToString().ToLower();
                                if (itemString.Contains(AudioBar2AppString.ToLower()))
                                {
                                    AudioBar2AppComboBox.SelectedItem = AudioBar2AppString.ToLower();
                                    found = true;
                                    break;
                                }
                            }

                            if (!found)
                            {
                                string notDetected = $"App Not Detected ( {AudioBar2AppString} )";
                                if (!AudioBar2AppComboBox.Items.Contains(notDetected))
                                    AudioBar2AppComboBox.Items.Insert(0, notDetected);

                                AudioBar2AppComboBox.SelectedIndex = 0;
                            }
                        }

                    });
                });
                await Task.Run(() =>
                {
                    Dispatcher.Invoke(() =>
                    {
                        if (!string.IsNullOrEmpty(AudioBar3AppString))
                        {
                            bool found = false;
                            foreach (var item in AudioBar3AppComboBox.Items)
                            {
                                var itemString = item.ToString().ToLower();
                                if (itemString.Contains(AudioBar3AppString.ToLower()))
                                {
                                    AudioBar3AppComboBox.SelectedItem = AudioBar3AppString.ToLower();
                                    found = true;
                                    break;
                                }
                            }

                            if (!found)
                            {
                                string notDetected = $"App Not Detected ( {AudioBar3AppString} )";
                                if (!AudioBar3AppComboBox.Items.Contains(notDetected))
                                    AudioBar3AppComboBox.Items.Insert(0, notDetected);

                                AudioBar3AppComboBox.SelectedIndex = 0;
                            }
                        }

                    });
                });
                await Task.Run(() =>
                {
                    Dispatcher.Invoke(() =>
                    {
                        if (!string.IsNullOrEmpty(AudioBar4AppString))
                        {
                            bool found = false;
                            foreach (var item in AudioBar4AppComboBox.Items)
                            {
                                var itemString = item.ToString().ToLower();
                                if (itemString.Contains(AudioBar4AppString.ToLower()))
                                {
                                    AudioBar4AppComboBox.SelectedItem = AudioBar4AppString.ToLower();
                                    found = true;
                                    break;
                                }
                            }

                            if (!found)
                            {
                                string notDetected = $"App Not Detected ( {AudioBar4AppString} )";
                                if (!AudioBar4AppComboBox.Items.Contains(notDetected))
                                    AudioBar4AppComboBox.Items.Insert(0, notDetected);

                                AudioBar4AppComboBox.SelectedIndex = 0;
                            }
                        }

                    });
                });
                await Task.Run(() =>
                {
                    Dispatcher.Invoke(() =>
                    {
                        if (!string.IsNullOrEmpty(AudioBar5AppString))
                        {
                            bool found = false;
                            foreach (var item in AudioBar5AppComboBox.Items)
                            {
                                var itemString = item.ToString().ToLower();
                                if (itemString.Contains(AudioBar5AppString.ToLower()))
                                {
                                    AudioBar5AppComboBox.SelectedItem = AudioBar5AppString.ToLower();
                                    found = true;
                                    break;
                                }
                            }

                            if (!found)
                            {
                                string notDetected = $"App Not Detected ( {AudioBar5AppString} )";
                                if (!AudioBar5AppComboBox.Items.Contains(notDetected))
                                    AudioBar5AppComboBox.Items.Insert(0, notDetected);

                                AudioBar5AppComboBox.SelectedIndex = 0;
                            }
                        }
                    });
                });

                if (DeviceVIDTextBox != null)
                {
                    DeviceVIDTextBox.Text = DeviceVIDString;
                }
                if (DevicePIDTextBox != null)
                {
                    DevicePIDTextBox.Text = DevicePIDString;
                }

                Debug.WriteLine("Settings Loaded Successfully!");
            }
            else
            {
                AudioBar1AppComboBox.Items.Insert(0, "Please Select Source");
                AudioBar1AppComboBox.SelectedIndex = 0;
                AudioBar2AppComboBox.Items.Insert(0, "Please Select Source");
                AudioBar2AppComboBox.SelectedIndex = 0;
                AudioBar3AppComboBox.Items.Insert(0, "Please Select Source");
                AudioBar3AppComboBox.SelectedIndex = 0;
                AudioBar4AppComboBox.Items.Insert(0, "Please Select Source");
                AudioBar4AppComboBox.SelectedIndex = 0;
                AudioBar5AppComboBox.Items.Insert(0, "Please Select Source");
                AudioBar5AppComboBox.SelectedIndex = 0;
                MessageBox.Show("Config file not found. Please create a config file.");
            }
        }

        // Populates the Combo Boxes with the currently running applications, has to be on its own MTA Thread
        private async Task PopulateComboBoxes()
        {
            Task.Run(() =>
            {
                var mtaThread = new Thread(() =>
                {
                    try
                    {
                        using (var enumerator = new MMDeviceEnumerator())
                        {
                            var defaultDevice = enumerator.GetDefaultAudioEndpoint(DataFlow.Render, Role.Console);
                            if (defaultDevice == null)
                            {
                                Debug.WriteLine("No default audio device found.");
                                return;
                            }

                            using (var sessionManager = AudioSessionManager2.FromMMDevice(defaultDevice))
                            {
                                var sessions = sessionManager.GetSessionEnumerator();
                                foreach (var session in sessions)
                                {
                                    using (var sessionControl = session.QueryInterface<AudioSessionControl2>())
                                    {
                                        if (sessionControl == null || sessionControl.Process == null)
                                            continue;

                                        string appName = sessionControl.Process.ProcessName;

                                        if (!string.IsNullOrEmpty(appName))
                                        {
                                            if (appName != "Idle")
                                            {
                                                Dispatcher.Invoke(() =>
                                                {
                                                    AudioBar1AppComboBox.Items.Add(appName.ToLower());
                                                    AudioBar2AppComboBox.Items.Add(appName.ToLower());
                                                    AudioBar3AppComboBox.Items.Add(appName.ToLower());
                                                    AudioBar4AppComboBox.Items.Add(appName.ToLower());
                                                    AudioBar5AppComboBox.Items.Add(appName.ToLower());
                                                });
                                            }
                                        }
                                    }
                                }
                            }
                        }
                        Debug.WriteLine("Combo boxes populated successfully.");
                    }
                    catch (Exception ex)
                    {
                        Debug.WriteLine($"Error populating combo boxes: {ex.Message}");
                    }
                });
                mtaThread.SetApartmentState(ApartmentState.MTA);
                mtaThread.Start();
            });
        }

        private void AudioBar1MuteButton_Click(object sender, RoutedEventArgs e)
        {
            if (cachedSessionControl1 != null)
            {
                var simpleAudioVolume = cachedSessionControl1.QueryInterface<SimpleAudioVolume>();
                if (simpleAudioVolume != null)
                {
                    simpleAudioVolume.IsMuted = !simpleAudioVolume.IsMuted;
                    AudioBar1MuteButton.Content = simpleAudioVolume.IsMuted ? "Unmute" : "Mute";
                }
            }
        }
        private void AudioBar2MuteButton_Click(object sender, RoutedEventArgs e)
        {
            if (cachedSessionControl2 != null)
            {
                var simpleAudioVolume = cachedSessionControl2.QueryInterface<SimpleAudioVolume>();
                if (simpleAudioVolume != null)
                {
                    simpleAudioVolume.IsMuted = !simpleAudioVolume.IsMuted;
                    AudioBar2MuteButton.Content = simpleAudioVolume.IsMuted ? "Unmute" : "Mute";
                }
            }
        }
        private void AudioBar3MuteButton_Click(object sender, RoutedEventArgs e)
        {
            if (cachedSessionControl3 != null)
            {
                var simpleAudioVolume = cachedSessionControl3.QueryInterface<SimpleAudioVolume>();
                if (simpleAudioVolume != null)
                {
                    simpleAudioVolume.IsMuted = !simpleAudioVolume.IsMuted;
                    AudioBar3MuteButton.Content = simpleAudioVolume.IsMuted ? "Unmute" : "Mute";
                }
            }
        }
        private void AudioBar4MuteButton_Click(object sender, RoutedEventArgs e)
        {
            if (cachedSessionControl4 != null)
            {
                var simpleAudioVolume = cachedSessionControl4.QueryInterface<SimpleAudioVolume>();
                if (simpleAudioVolume != null)
                {
                    simpleAudioVolume.IsMuted = !simpleAudioVolume.IsMuted;
                    AudioBar4MuteButton.Content = simpleAudioVolume.IsMuted ? "Unmute" : "Mute";
                }
            }
        }
        private void AudioBar5MuteButton_Click(object sender, RoutedEventArgs e)
        {
            if (cachedSessionControl5 != null)
            {
                var simpleAudioVolume = cachedSessionControl5.QueryInterface<SimpleAudioVolume>();
                if (simpleAudioVolume != null)
                {
                    simpleAudioVolume.IsMuted = !simpleAudioVolume.IsMuted;
                    AudioBar5MuteButton.Content = simpleAudioVolume.IsMuted ? "Unmute" : "Mute";
                }
            }
        }

        private void SaveButton_Click(object sender, RoutedEventArgs e) // Saves config to file in the same directory as the executable
        {
            string filePath = System.IO.Path.Combine(Environment.CurrentDirectory, "config.json");

            string DeviceVIDSaveItem = DeviceVIDTextBox.Text;
            string DevicePIDSaveItem = DevicePIDTextBox.Text;

            string AudioBar1AppSaveItem = AudioBar1AppComboBox.SelectedItem.ToString().ToLower();
            string AudioBar2AppSaveItem = AudioBar2AppComboBox.SelectedItem.ToString().ToLower();
            string AudioBar3AppSaveItem = AudioBar3AppComboBox.SelectedItem.ToString().ToLower();
            string AudioBar4AppSaveItem = AudioBar4AppComboBox.SelectedItem.ToString().ToLower();
            string AudioBar5AppSaveItem = AudioBar5AppComboBox.SelectedItem.ToString().ToLower();

            if (AudioBar1AppSaveItem.StartsWith("app not detected ( ") && AudioBar1AppSaveItem.EndsWith(" )"))
            {
                int startIndex = AudioBar1AppSaveItem.IndexOf("( ") + 1;
                int endIndex = AudioBar1AppSaveItem.IndexOf(" )");
                AudioBar1AppSaveItem = AudioBar1AppSaveItem.Substring(startIndex, endIndex - startIndex).Trim();
            }
            if (AudioBar2AppSaveItem.StartsWith("app not detected ( ") && AudioBar2AppSaveItem.EndsWith(" )"))
            {
                int startIndex = AudioBar2AppSaveItem.IndexOf("( ") + 1;
                int endIndex = AudioBar2AppSaveItem.IndexOf(" )");
                AudioBar2AppSaveItem = AudioBar2AppSaveItem.Substring(startIndex, endIndex - startIndex).Trim();
            }
            if (AudioBar3AppSaveItem.StartsWith("app not detected ( ") && AudioBar3AppSaveItem.EndsWith(" )"))
            {
                int startIndex = AudioBar3AppSaveItem.IndexOf("( ") + 1;
                int endIndex = AudioBar3AppSaveItem.IndexOf(" )");
                AudioBar3AppSaveItem = AudioBar3AppSaveItem.Substring(startIndex, endIndex - startIndex).Trim();
            }
            if (AudioBar4AppSaveItem.StartsWith("app not detected ( ") && AudioBar4AppSaveItem.EndsWith(" )"))
            {
                int startIndex = AudioBar4AppSaveItem.IndexOf("( ") + 1;
                int endIndex = AudioBar4AppSaveItem.IndexOf(" )");
                AudioBar4AppSaveItem = AudioBar4AppSaveItem.Substring(startIndex, endIndex - startIndex).Trim();
            }
            if (AudioBar5AppSaveItem.StartsWith("app not detected ( ") && AudioBar5AppSaveItem.EndsWith(" )"))
            {
                int startIndex = AudioBar5AppSaveItem.IndexOf("( ") + 1;
                int endIndex = AudioBar5AppSaveItem.IndexOf(" )");
                AudioBar5AppSaveItem = AudioBar5AppSaveItem.Substring(startIndex, endIndex - startIndex).Trim();
            }

            var settings = new JObject()
            {
                ["appSource"] = new JObject
                {
                    ["AudioBar1"] = AudioBar1AppSaveItem ?? "Not Set",
                    ["AudioBar2"] = AudioBar2AppSaveItem ?? "Not Set",
                    ["AudioBar3"] = AudioBar3AppSaveItem ?? "Not Set",
                    ["AudioBar4"] = AudioBar4AppSaveItem ?? "Not Set",
                    ["AudioBar5"] = AudioBar5AppSaveItem ?? "Not Set",
                },
                ["device"] = new JObject
                {
                    ["DeviceVID"] = DeviceVIDSaveItem ?? "Not Set",
                    ["DevicePID"] = DevicePIDSaveItem ?? "Not Set",
                },
            };

            string settingsJson = settings.ToString();
            System.IO.File.WriteAllText(filePath, settingsJson);
            Debug.WriteLine("Settings saved successfully.");

            Debug.WriteLine("Restarting...");
            if (_audioBarTimer != null)
            {
                _audioBarTimer.Stop();
                _audioBarTimer = null;
            }
            AudioBar1AppComboBox.Items.Clear();
            AudioBar2AppComboBox.Items.Clear();
            AudioBar3AppComboBox.Items.Clear();
            AudioBar4AppComboBox.Items.Clear();
            AudioBar5AppComboBox.Items.Clear();
            AsyncLoading();
        }

        private void GetAudioSession(string appName, Action<AudioMeterInformation> setSessionCallback, string audioBarId)
        {
            var audioSessionThread = new Thread(() =>
            {
                try
                {
                    using (var deviceEnumerator = new MMDeviceEnumerator())
                    {
                        var defaultDevice = deviceEnumerator.GetDefaultAudioEndpoint(DataFlow.Render, Role.Multimedia);
                        if (defaultDevice == null)
                        {
                            Debug.WriteLine("No default audio device found.");
                            return;
                        }

                        using (var sessionManager = AudioSessionManager2.FromMMDevice(defaultDevice))
                        {
                            var sessions = sessionManager.GetSessionEnumerator();
                            foreach (var session in sessions)
                            {
                                var session2 = session.QueryInterface<AudioSessionControl2>();
                                if (session2 != null && session2.Process != null)
                                {
                                    string processName = session2.Process.ProcessName.ToLower();
                                    bool match = processName.Contains(appName.ToLower());
                                    if (match)
                                    {
                                        var audioMeterInfo = session.QueryInterface<AudioMeterInformation>();
                                        if (audioMeterInfo != null)
                                        {
                                            setSessionCallback(audioMeterInfo);
                                        }
                                        else
                                        {
                                            Debug.WriteLine($"AudioMeterInformation is NULL for {appName}.");
                                        }
                                        return;
                                    }

                                    if (session2.Process.ProcessName.ToLower().Contains(appName))
                                    {
                                        var audioMeterInfo = session.QueryInterface<AudioMeterInformation>();
                                        if (audioMeterInfo != null)
                                        {
                                            setSessionCallback(audioMeterInfo);
                                        }
                                        else
                                        {
                                            Debug.WriteLine($"AudioMeterInformation is NULL for {appName}.");
                                        }
                                        return;
                                    }
                                }
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    Debug.WriteLine($"Error getting audio session: {ex.Message}");
                }
            });
            audioSessionThread.Start();
        }

        // Starts listening to the Audio Session Volume and updates the Audio Bar UI
        private void StartAudioMonitorThread(string appName, Rectangle audioBar, Func<AudioMeterInformation> getSession)
        {
            var mtaThread = new Thread(() =>
            {
                try
                {
                    while (true)
                    {
                        try
                        {
                            var targetSession = getSession();
                            if (targetSession == null)
                            {
                                Dispatcher.Invoke(() =>
                                {
                                    audioBar.Height = 300;
                                });
                            }
                            else
                            {
                                float peakValue = targetSession.PeakValue;
                                float newValue = peakValue * -100 + 100;
                                newValue = newValue * 3;
                                if (newValue < 0) { newValue = 0; }
                                Dispatcher.Invoke(() =>
                                {
                                    audioBar.Height = newValue;
                                });

                            }
                        }
                        catch (Exception e)
                        {
                            Debug.WriteLine($"Error getting audio level for {appName}: {e.Message}");
                        }
                        Thread.Sleep(10);
                    }
                }
                catch (Exception ex)
                {
                    Debug.WriteLine($"Error monitoring audio level for {appName}: {ex.Message}");
                }
            });
            mtaThread.SetApartmentState(ApartmentState.MTA);
            mtaThread.Start();
        }

        // Sets the Audio Session/App and updates the Audio Bar Name Label
        private async Task GetAudioBarValueSession(string appName, int audioBarNumber)
        {
            var audioSessionThread = new Thread(() =>
            {
                try
                {
                    using (var deviceEnumerator = new MMDeviceEnumerator())
                    {
                        var defaultDevice = deviceEnumerator.GetDefaultAudioEndpoint(DataFlow.Render, Role.Multimedia);
                        if (defaultDevice == null)
                        {
                            Debug.WriteLine("No default audio device found.");
                            return;
                        }

                        using (var sessionManager = AudioSessionManager2.FromMMDevice(defaultDevice))
                        {
                            var sessions = sessionManager.GetSessionEnumerator();
                            foreach (var session in sessions)
                            {
                                var session2 = session.QueryInterface<AudioSessionControl2>();
                                if (session2 != null && session2.Process != null)
                                {
                                    string processName = session2.Process.ProcessName.ToLower();
                                    bool match = processName.Contains(appName.ToLower());

                                    if (match)
                                    {
                                        switch (audioBarNumber)
                                        {
                                            case 1:
                                                cachedSessionControl1 = session2;
                                                break;
                                            case 2:
                                                cachedSessionControl2 = session2;
                                                break;
                                            case 3:
                                                cachedSessionControl3 = session2;
                                                break;
                                            case 4:
                                                cachedSessionControl4 = session2;
                                                break;
                                            case 5:
                                                cachedSessionControl5 = session2;
                                                break;
                                            default:
                                                Debug.WriteLine("Invalid audio bar number.");
                                                return;
                                        }
                                    }
                                    if (session2.Process.ProcessName.ToLower().Contains(appName))
                                    {
                                        switch (audioBarNumber)
                                        {
                                            case 1:
                                                cachedSessionControl1 = session2;
                                                break;
                                            case 2:
                                                cachedSessionControl2 = session2;
                                                break;
                                            case 3:
                                                cachedSessionControl3 = session2;
                                                break;
                                            case 4:
                                                cachedSessionControl4 = session2;
                                                break;
                                            case 5:
                                                cachedSessionControl5 = session2;
                                                break;
                                            default:
                                                Debug.WriteLine("Invalid audio bar number.");
                                                return;
                                        }
                                        return;
                                    }
                                }
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    Debug.WriteLine($"Error getting audio session: {ex.Message}");
                }
            });
            audioSessionThread.Start();
        }

        // Reads the Audio Ouput Level for the selected Audio Session/App and updates the UI
        private async Task GetAudioBarValue(int audioBarNumber, string appName, Label audioBarLabel, Button audioBarMuteButton)
        {
            AudioSessionControl2 sessionControl = null;
            Rectangle sessionMuteButtonRectangle = null;
            Label sessionMuteButtonLabel = null;

            switch (audioBarNumber)
            {
                case 1:
                    sessionControl = cachedSessionControl1;
                    break;
                case 2:
                    sessionControl = cachedSessionControl2;
                    break;
                case 3:
                    sessionControl = cachedSessionControl3;
                    break;
                case 4:
                    sessionControl = cachedSessionControl4;
                    break;
                case 5:
                    sessionControl = cachedSessionControl5;
                    break;
                default:
                    Debug.WriteLine("Invalid audio bar number.");
                    return;
            }

            if (sessionControl != null)
            {
                var simpleAudioVolume = sessionControl.QueryInterface<SimpleAudioVolume>();
                if (simpleAudioVolume != null)
                {
                    float volume = simpleAudioVolume.MasterVolume * 100;
                    int intVolume = (int)volume;
                    bool muted = simpleAudioVolume.IsMuted;

                    await Dispatcher.InvokeAsync(() =>
                    {
                        if (audioBarLabel != null)
                        {
                            if (!string.IsNullOrEmpty(appName))
                            {
                                string capitalizedAppName = char.ToUpper(appName[0]) + appName.Substring(1);
                            }

                            if (muted)
                            {
                                audioBarMuteButton.Content = "Muted";
                                audioBarMuteButton.Background = Brushes.Red;
                                audioBarMuteButton.Foreground = Brushes.White;
                                audioBarLabel.Content = "Muted";
                                audioBarLabel.Foreground = Brushes.Red;
                            }
                            else
                            {
                                audioBarMuteButton.Content = "Mute";
                                audioBarMuteButton.Background = Brushes.LightGray;
                                audioBarMuteButton.Foreground = Brushes.Black;
                                audioBarLabel.Foreground = Brushes.Black;
                                audioBarLabel.Content = $"{intVolume}%";
                            }
                        }
                        else
                        {
                            if (audioBarLabel != null)
                            {
                                audioBarLabel.Content = "Closed";
                            }

                            if (!string.IsNullOrEmpty(appName))
                            {
                                string capitalizedAppName = char.ToUpper(appName[0]) + appName.Substring(1);
                            }
                        }
                    });
                }
                else
                {
                    Debug.WriteLine($"Failed to retrieve SimpleAudioVolume for {appName}.");
                }
            }
            else
            {
                await Dispatcher.InvokeAsync(() =>
                {
                    if (audioBarLabel != null)
                    {
                        audioBarLabel.Content = "100%";
                    }

                    if (!string.IsNullOrEmpty(appName))
                    {
                        string capitalizedAppName = char.ToUpper(appName[0]) + appName.Substring(1);
                    }
                });
            }
        }

        private async void AudioBarTimer_Tick(object sender, EventArgs e)
        {
            GetAudioBarValue(1, AudioBar1AppComboBox.SelectedItem.ToString(), AudioBar1VolumeLabel, AudioBar1MuteButton);
            GetAudioBarValue(2, AudioBar2AppComboBox.SelectedItem.ToString(), AudioBar2VolumeLabel, AudioBar2MuteButton);
            GetAudioBarValue(3, AudioBar3AppComboBox.SelectedItem.ToString(), AudioBar3VolumeLabel, AudioBar3MuteButton);
            GetAudioBarValue(4, AudioBar4AppComboBox.SelectedItem.ToString(), AudioBar4VolumeLabel, AudioBar4MuteButton);
            GetAudioBarValue(5, AudioBar5AppComboBox.SelectedItem.ToString(), AudioBar5VolumeLabel, AudioBar5MuteButton);
        }

        private async void RestartBackgroundTaskButton_Click(object sender, RoutedEventArgs e)
        {
            string backgroundTaskPath = System.IO.Path.Combine(Environment.CurrentDirectory, "Smart Audio Mixer Background Task.exe");
            if (!File.Exists(backgroundTaskPath))
            {
                MessageBoxResult result = MessageBox.Show("Background Task Executable Not Found! \n Please Redownload The Application", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                Process.Start(new ProcessStartInfo("https://github.com/TeagueGillard/Smart-Audio-Mixer/releases") { UseShellExecute = true });
                return;
            }
            bool BackgroundTaskRunning = Process.GetProcessesByName("Smart Audio Mixer Background Task").Any();
            if (BackgroundTaskRunning)
            {
                Process.Start("taskkill", "/F /IM Smart Audio Mixer Background Task.exe");
                await Task.Delay(1000);
                Process.Start(backgroundTaskPath);
            }
            else
            {
                Process.Start(backgroundTaskPath);
            }
        }

        private void HardwareScanTimer_Tick(object sender, EventArgs e)
        {
            string comPort = null;
            string searchPattern = $"VID_{DeviceVIDTextBox.Text}&PID_{DevicePIDTextBox.Text}";
            bool deviceConnected = false;

            RegistryKey usbDevicesKey = Registry.LocalMachine.OpenSubKey(@"SYSTEM\CurrentControlSet\Enum\USB");

            if (usbDevicesKey != null)
            {
                foreach (string subKeyName in usbDevicesKey.GetSubKeyNames())
                {
                    if (subKeyName.Contains(searchPattern))
                    {
                        RegistryKey deviceKey = usbDevicesKey.OpenSubKey(subKeyName);
                        if (deviceKey != null)
                        {
                            foreach (string deviceInstance in deviceKey.GetSubKeyNames())
                            {
                                RegistryKey instanceKey = deviceKey.OpenSubKey(deviceInstance);
                                if (instanceKey != null)
                                {
                                    RegistryKey deviceParamsKey = instanceKey.OpenSubKey("Device Parameters");
                                    if (deviceParamsKey != null)
                                    {
                                        comPort = deviceParamsKey.GetValue("PortName") as string;
                                        if (!string.IsNullOrEmpty(comPort))
                                        {
                                            string[] SerialPorts = SerialPort.GetPortNames();
                                            foreach (string port in SerialPorts)
                                            {
                                                if (port == comPort)
                                                {
                                                    deviceConnected = true;
                                                }
                                                else
                                                {
                                                    deviceConnected = false;
                                                }
                                            }
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
            }

            if (deviceConnected)
            {
                DeviceStatusLabel.Content = "Connected";
                DeviceStatusLabel.Foreground = Brushes.Green;
            }
            else
            {
                DeviceStatusLabel.Content = "Not Connected";
                DeviceStatusLabel.Foreground = Brushes.Red;
            }
        }

        private void BackgroundTaskScan_Tick(object sender, EventArgs e)
        {
            bool BackgroundTaskRunning = Process.GetProcessesByName("Smart Audio Mixer Background Task").Any();
            if (BackgroundTaskRunning)
            {
                BackgroundTaskStatusLabel.Content = "Running";
                BackgroundTaskStatusLabel.Foreground = Brushes.Green;
                RestartBackgroundTaskButton.Content = "Restart Background Task";
            }
            else
            {
                BackgroundTaskStatusLabel.Content = "Not Running";
                BackgroundTaskStatusLabel.Foreground = Brushes.Red;
                RestartBackgroundTaskButton.Content = "Start Background Task";
            }
        }

        protected override void OnClosing(System.ComponentModel.CancelEventArgs e)
        {
            if (_audioBarTimer != null)
            {
                _audioBarTimer.Stop();
                _audioBarTimer = null;
            }
            Application.Current.Shutdown();
            Process.GetCurrentProcess().Kill();
            Environment.Exit(0);
        }
    }
}