using Microsoft.Win32;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System.Diagnostics;
using System.IO;
using System.IO.Ports;
using NAudio.CoreAudioApi;

class Program
{
    static string filePath = System.IO.Path.Combine(Environment.CurrentDirectory, "config.json"); // Looks for config.json in the same directory as the executable

    static string AudioBar1AppString = null;
    static string AudioBar2AppString = null;
    static string AudioBar3AppString = null;
    static string AudioBar4AppString = null;
    static string AudioBar5AppString = null;

    static SerialPort _serialPort;

    static readonly MMDeviceEnumerator deviceEnumerator = new();
    static readonly MMDevice audioDevice = deviceEnumerator.GetDefaultAudioEndpoint(DataFlow.Render, Role.Multimedia);
    static readonly Dictionary<string, SimpleAudioVolume> sessionVolumeCache = new(StringComparer.OrdinalIgnoreCase);

    static void Main()
    {
        if (File.Exists(filePath))
        {
            #if DEBUG
            Debug.WriteLine("Config file found. Reading settings...");
            #endif

            JObject settings = JObject.Parse(System.IO.File.ReadAllText(filePath));

            AudioBar1AppString = (string)settings["appSource"]["AudioBar1"];
            AudioBar2AppString = (string)settings["appSource"]["AudioBar2"];
            AudioBar3AppString = (string)settings["appSource"]["AudioBar3"];
            AudioBar4AppString = (string)settings["appSource"]["AudioBar4"];
            AudioBar5AppString = (string)settings["appSource"]["AudioBar5"];

            string DeviceVIDString = (string)settings["device"]["DeviceVID"];
            string DevicePIDString = (string)settings["device"]["DevicePID"];

            #if DEBUG
            Debug.WriteLine("Loaded Settings...");
            #endif

            string comPort = null;
            string searchPattern = $"VID_{DeviceVIDString}&PID_{DevicePIDString}";
            bool comPortFound = false;

            RegistryKey usbDevicesKey = Registry.LocalMachine.OpenSubKey(@"SYSTEM\CurrentControlSet\Enum\USB");

            #if DEBUG
            Debug.WriteLine("Searching for Device...");
            #endif

            if (usbDevicesKey != null)
            {
                foreach (string subKeyName in usbDevicesKey.GetSubKeyNames())
                {
                    if (subKeyName.Contains(searchPattern, StringComparison.OrdinalIgnoreCase))
                    {
                        RegistryKey deviceKey = usbDevicesKey.OpenSubKey(subKeyName);
                        if (deviceKey != null)
                        {
                            foreach (string instanceName in deviceKey.GetSubKeyNames())
                            {
                                RegistryKey instanceKey = deviceKey.OpenSubKey(instanceName);
                                RegistryKey parametersKey = instanceKey?.OpenSubKey("Device Parameters");
                                
                                #if DEBUG
                                Debug.WriteLine($"Found Device: {instanceName}");
                                #endif

                                if (parametersKey != null)
                                {
                                    string portName = parametersKey.GetValue("PortName") as string;
                                    if (!string.IsNullOrWhiteSpace(portName))
                                    {
                                        comPort = portName;
                                        break;
                                    }
                                }
                            }
                        }
                    }
                    if (!string.IsNullOrEmpty(comPort))
                        break;
                }
            }
            if (!string.IsNullOrWhiteSpace(comPort))
            {
                #if DEBUG
                Debug.WriteLine($"Found COM port: {comPort}");
                #endif

                _serialPort = new SerialPort(comPort, 115200);
                _serialPort.ReadTimeout = 500;
                _serialPort.WriteTimeout = 500;
                _serialPort.DtrEnable = true;
                _serialPort.RtsEnable = true;
                _serialPort.DataReceived += SerialDataReceived;
                _serialPort.Open();
                while (true)
                {
                    Thread.Sleep(50);
                }
            }
            else
            {
                Console.WriteLine("COM port not found.");
            }

        }
    }

    static void SerialDataReceived(object sender, SerialDataReceivedEventArgs e)
    {
        #if DEBUG
        Debug.WriteLine("Data Received");
        #endif

        try
        {
            string line = _serialPort.ReadLine().Trim();
            if (line.Length > 120)
            {
                _serialPort.DiscardInBuffer();
                return;
            }

            #if DEBUG
            Debug.WriteLine($"Received: {line}");
            #endif

            int v1 = 0, v2 = 0, v3 = 0, v4 = 0, v5 = 0;
            int m1 = 0, m2 = 0, m3 = 0, m4 = 0, m5 = 0;

            string[] pairs = line.Split(',');

            foreach (var pair in pairs)
            {
                int colonIndex = pair.IndexOf(':');
                if (colonIndex <= 0 || colonIndex == pair.Length - 1)
                    continue;

                string key = pair.Substring(0, colonIndex).Trim();
                string valStr = pair.Substring(colonIndex + 1).Trim();

                if (!int.TryParse(valStr, out int val))
                    continue;

                switch (key)
                {
                    case "Volume1": v1 = val; break;
                    case "Volume2": v2 = val; break;
                    case "Volume3": v3 = val; break;
                    case "Volume4": v4 = val; break;
                    case "Volume5": v5 = val; break;
                    case "Mute1": m1 = val; break;
                    case "Mute2": m2 = val; break;
                    case "Mute3": m3 = val; break;
                    case "Mute4": m4 = val; break;
                    case "Mute5": m5 = val; break;
                }
            }

            SetAppVolumeAndMute(AudioBar1AppString, v1 / 100f, m1 == 1);
            SetAppVolumeAndMute(AudioBar2AppString, v2 / 100f, m2 == 1);
            SetAppVolumeAndMute(AudioBar3AppString, v3 / 100f, m3 == 1);
            SetAppVolumeAndMute(AudioBar4AppString, v4 / 100f, m4 == 1);
            SetAppVolumeAndMute(AudioBar5AppString, v5 / 100f, m5 == 1);

            GC.Collect();
            GC.WaitForPendingFinalizers();
            GC.Collect();
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Serial error: {ex.Message}");
        }
    }

    static void SetAppVolumeAndMute(string appName, float volume, bool mute)
    {
        if (string.IsNullOrWhiteSpace(appName)) return;

        try
        {
            var sessions = audioDevice.AudioSessionManager.Sessions;

            for (int i = 0; i < sessions.Count; i++)
            {
                using var session = sessions[i];

                if (session.GetSessionIdentifier.Contains(appName, StringComparison.OrdinalIgnoreCase))
                {
                    var simpleVolume = session.SimpleAudioVolume;
                    simpleVolume.Volume = volume;
                    simpleVolume.Mute = mute;

                    #if DEBUG
                    Debug.WriteLine($"Set {appName}: Volume={volume * 100:F0}, Mute={mute}");
                    #endif

                    break;
                }
            }
        }
        catch (Exception ex)
        {
            #if DEBUG
            Debug.WriteLine($"Error setting volume for {appName}: {ex.Message}");
            #endif
        }
    }
}