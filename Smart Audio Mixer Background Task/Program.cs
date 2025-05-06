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

    static void Main()
    {
        if (File.Exists(filePath))
        {
            Debug.WriteLine("Config file found. Reading settings...");
            JObject settings = JObject.Parse(System.IO.File.ReadAllText(filePath));

            AudioBar1AppString = (string)settings["appSource"]["AudioBar1"];
            AudioBar2AppString = (string)settings["appSource"]["AudioBar2"];
            AudioBar3AppString = (string)settings["appSource"]["AudioBar3"];
            AudioBar4AppString = (string)settings["appSource"]["AudioBar4"];
            AudioBar5AppString = (string)settings["appSource"]["AudioBar5"];

            string DeviceVIDString = (string)settings["device"]["DeviceVID"];
            string DevicePIDString = (string)settings["device"]["DevicePID"];

            Debug.WriteLine("Loaded Settings...");

            string comPort = null;
            string searchPattern = $"VID_{DeviceVIDString}&PID_{DevicePIDString}";
            bool comPortFound = false;

            RegistryKey usbDevicesKey = Registry.LocalMachine.OpenSubKey(@"SYSTEM\CurrentControlSet\Enum\USB");

            Debug.WriteLine("Searching for Device...");
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
                                Debug.WriteLine($"Found Device: {instanceName}");
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
                Debug.WriteLine($"Found COM port: {comPort}");
                _serialPort = new SerialPort(comPort, 115200);
                _serialPort.ReadTimeout = 500;
                _serialPort.WriteTimeout = 500;
                _serialPort.DtrEnable = true;
                _serialPort.RtsEnable = true;
                _serialPort.DataReceived += SerialDataReceived;
                _serialPort.Open();
                while (true)
                {
                    Thread.Sleep(300);
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
        Debug.WriteLine("Data Recieved");
        try
        {
            string line = _serialPort.ReadLine().Trim();
            Debug.WriteLine($"Received: {line}");

            var values = line
                .Split(',', StringSplitOptions.RemoveEmptyEntries)
                .Select(s => s.Trim().Split(':'))
                .Where(p => p.Length == 2 && int.TryParse(p[1], out _))
                .ToDictionary(p => p[0], p => int.Parse(p[1]));

            float Volume1 = values.TryGetValue("Volume1", out int v1) ? v1 / 100f : 0f;
            float Volume2 = values.TryGetValue("Volume2", out int v2) ? v2 / 100f : 0f;
            float Volume3 = values.TryGetValue("Volume3", out int v3) ? v3 / 100f : 0f;
            float Volume4 = values.TryGetValue("Volume4", out int v4) ? v4 / 100f : 0f;
            float Volume5 = values.TryGetValue("Volume5", out int v5) ? v5 / 100f : 0f;

            bool isMuted1 = values.TryGetValue("Mute1", out int m1) && m1 == 1;
            bool isMuted2 = values.TryGetValue("Mute2", out int m2) && m2 == 1;
            bool isMuted3 = values.TryGetValue("Mute3", out int m3) && m3 == 1;
            bool isMuted4 = values.TryGetValue("Mute4", out int m4) && m4 == 1;
            bool isMuted5 = values.TryGetValue("Mute5", out int m5) && m5 == 1;

            SetAppVolumeAndMute(AudioBar1AppString, Volume1, isMuted1);
            SetAppVolumeAndMute(AudioBar2AppString, Volume2, isMuted2);
            SetAppVolumeAndMute(AudioBar3AppString, Volume3, isMuted3);
            SetAppVolumeAndMute(AudioBar4AppString, Volume4, isMuted4);
            SetAppVolumeAndMute(AudioBar5AppString, Volume5, isMuted5);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Serial error: {ex.Message}");
        }
    }

    static void SetAppVolumeAndMute(string appName, float volume, bool mute)
    {
        Debug.WriteLine($"Setting volume for {appName}: Volume={volume * 100:F0}, Mute={mute}");
        var enumerator = new MMDeviceEnumerator();
        var device = enumerator.GetDefaultAudioEndpoint(DataFlow.Render, Role.Multimedia);
        var sessions = device.AudioSessionManager.Sessions;

        for (int i = 0; i < sessions.Count; i++)
        {
            var session = sessions[i];
            try
            {
                if (session.GetSessionIdentifier.Contains(appName, StringComparison.OrdinalIgnoreCase))
                {
                    session.SimpleAudioVolume.Volume = volume;
                    session.SimpleAudioVolume.Mute = mute;
                    Debug.WriteLine($"Set {session.DisplayName}: Volume={volume * 100:F0}, Mute={mute}");
                    session.Dispose();
                    session = null;
                    sessions = null;
                    device = null;
                    enumerator = null;
                }
            }
            catch { }
        }
    }
}