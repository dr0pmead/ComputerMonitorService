﻿using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Management;
using System.Net;
using System.Net.Http;
using System.Net.NetworkInformation;
using System.Net.Sockets;
using System.ServiceProcess;
using System.Threading;
using System.Threading.Tasks;
using Newtonsoft.Json;

namespace ComputerMonitorService
{
    public partial class Service1 : ServiceBase
    {
        private string configFilePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "config.json");
        private HttpClient httpClient;
        private CancellationTokenSource cancellationTokenSource;
        private Task sendComputerInfoTask;
        private string serverUrl;

        public Service1()
        {
            InitializeComponent();
            try
            {
                httpClient = new HttpClient();
                httpClient.DefaultRequestHeaders.UserAgent.ParseAdd("Mozilla/5.0 (compatible; AcmeInc/1.0)");
                Log("HTTP client initialized successfully.");

                serverUrl = ReadServerUrlFromConfig();
                Log($"Server URL set to: {serverUrl}");
            }
            catch (Exception ex)
            {
                Log($"Error initializing service: {ex.Message}");
                throw;
            }
        }

        protected override void OnStart(string[] args)
        {
            Log("Service starting...");
            try
            {
                cancellationTokenSource = new CancellationTokenSource();
                var token = cancellationTokenSource.Token;

                sendComputerInfoTask = Task.Run(() => SendComputerInfo(token), token);

                Log("Service started successfully.");
            }
            catch (Exception ex)
            {
                Log($"Error during service start: {ex.Message}");
                throw;
            }
        }

        protected override void OnStop()
        {
            Log("Service stopping...");
            try
            {
                cancellationTokenSource.Cancel();
                if (sendComputerInfoTask != null)
                {
                    sendComputerInfoTask.Wait(TimeSpan.FromSeconds(10));  // Добавлено ожидание завершения задачи с таймаутом
                }
                Log("Service stopped successfully.");
            }
            catch (Exception ex)
            {
                Log($"Error during service stop: {ex.Message}");
                throw;
            }
        }

        private async Task SendComputerInfo(CancellationToken token)
        {
            while (!token.IsCancellationRequested)
            {
                try
                {
                    Log("Starting to gather computer information...");
                    var configData = ReadConfigFile();
                    var computerInfo = new
                    {
                        name = Environment.MachineName,
                        ipAddress = GetLocalIPAddresses(),
                        components = GetHardwareInfo(),
                        anyDesk = GetAnyDeskId(),
                        teamViewer = GetTeamViewerId(),
                        printer = GetDefaultPrinter(),
                        online = true,
                        owner = configData.owner,
                        department = configData.department,
                        lastUpdated = DateTime.UtcNow
                    };

                    string statusUrl = $"https://{serverUrl}/api/status";
                    Log(statusUrl);
                    Log("Sending data to server...");

                    // Convert computerInfo to JSON and log it
                    string computerInfoJson = JsonConvert.SerializeObject(computerInfo, Formatting.Indented);
                    Log("Collected Computer Info:");
                    Log(computerInfoJson);

                    var content = new StringContent(computerInfoJson, System.Text.Encoding.UTF8, "application/json");
                    var response = await httpClient.PostAsync(statusUrl, content);
                    response.EnsureSuccessStatusCode();
                    Log($"Data sent successfully. IPs: {string.Join(", ", computerInfo.ipAddress)}");

                    Log("Waiting for next cycle...");
                    await Task.Delay(TimeSpan.FromSeconds(60), token); // Увеличено время ожидания до 60 секунд
                }
                catch (Exception ex)
                {
                    Log($"Error in SendComputerInfo: {ex.Message}");
                }
            }
        }

        private dynamic ReadConfigFile()
        {
            try
            {
                Log("Reading config file...");
                var json = File.ReadAllText(configFilePath);
                return JsonConvert.DeserializeObject<dynamic>(json);
            }
            catch (Exception ex)
            {
                Log($"Error reading config file: {ex.Message}");
                return new { owner = "unknown", department = "unknown" };
            }
        }

        private string ReadServerUrlFromConfig()
        {
            try
            {
                Log("Reading server URL from config file...");
                var json = File.ReadAllText(configFilePath);
                dynamic config = JsonConvert.DeserializeObject<dynamic>(json);
                return config.serverUrl;
            }
            catch (Exception ex)
            {
                Log($"Error reading server URL from config file: {ex.Message}");
                throw;
            }
        }

        private string[] GetLocalIPAddresses()
        {
            try
            {
                Log("Getting local IP addresses...");
                var host = Dns.GetHostEntry(Dns.GetHostName());
                var ipList = new List<string>();

                foreach (var ip in host.AddressList)
                {
                    if (ip.AddressFamily == AddressFamily.InterNetwork)
                    {
                        ipList.Add(ip.ToString());
                    }
                }

                if (ipList.Count == 0)
                {
                    throw new Exception("Local IP Addresses Not Found!");
                }

                return ipList.ToArray();
            }
            catch (Exception ex)
            {
                Log($"Error getting local IP addresses: {ex.Message}");
                throw;
            }
        }

        private dynamic GetHardwareInfo()
        {
            var components = new List<dynamic>();

            try
            {
                Log("Getting hardware info...");
                var searcher = new ManagementObjectSearcher("select * from Win32_Processor");
                foreach (var obj in searcher.Get())
                {
                    components.Add(new
                    {
                        Type = "Processor",
                        Name = obj["Name"].ToString()
                    });
                }

                searcher = new ManagementObjectSearcher("select * from Win32_PhysicalMemory");
                var memoryInfo = new Dictionary<string, ulong>();
                foreach (var obj in searcher.Get())
                {
                    ulong capacity = Convert.ToUInt64(obj["Capacity"]);
                    string manufacturer = obj["Manufacturer"].ToString();

                    if (memoryInfo.ContainsKey(manufacturer))
                    {
                        memoryInfo[manufacturer] += capacity;
                    }
                    else
                    {
                        memoryInfo[manufacturer] = capacity;
                    }
                }

                foreach (var item in memoryInfo)
                {
                    components.Add(new
                    {
                        Type = "Memory",
                        Manufacturer = item.Key,
                        Quantity = (double)item.Value / (1024 * 1024 * 1024) // Конвертируем в гигабайты
                    });
                }

                searcher = new ManagementObjectSearcher("select * from Win32_DiskDrive");
                var diskToPartitions = new Dictionary<string, List<string>>();

                foreach (var obj in searcher.Get())
                {
                    var diskName = obj["Model"].ToString();
                    var diskSize = Convert.ToDouble(obj["Size"]);
                    var partitions = GetPartitionsForDisk(obj["DeviceID"].ToString());

                    diskToPartitions[diskName] = partitions;

                    components.Add(new
                    {
                        Type = "Disk",
                        Name = diskName,
                        Size = diskSize / (1024 * 1024 * 1024), // Конвертируем в гигабайты
                        FreeSpace = GetTotalFreeSpace(partitions) / (1024 * 1024 * 1024) // Конвертируем в гигабайты
                    });
                }

                searcher = new ManagementObjectSearcher("select * from Win32_VideoController");
                foreach (var obj in searcher.Get())
                {
                    var videoName = obj["Name"].ToString();
                    if (!videoName.Contains("Intel") && !videoName.Contains("Microsoft Basic Display Adapter") && !videoName.Contains("Radeon"))
                    {
                        components.Add(new
                        {
                            Type = "Video",
                            Name = videoName
                        });
                    }
                }

                searcher = new ManagementObjectSearcher("select * from Win32_BaseBoard");
                foreach (var obj in searcher.Get())
                {
                    components.Add(new
                    {
                        Type = "Motherboard",
                        Name = obj["Product"].ToString()
                    });
                }

                return components;
            }
            catch (Exception ex)
            {
                Log($"Error getting hardware info: {ex.Message}");
                throw;
            }
        }

        private List<string> GetPartitionsForDisk(string deviceId)
        {
            var partitions = new List<string>();

            var searcher = new ManagementObjectSearcher($"associators of {{Win32_DiskDrive.DeviceID='{deviceId}'}} where AssocClass=Win32_DiskDriveToDiskPartition");
            foreach (var obj in searcher.Get())
            {
                partitions.Add(obj["DeviceID"].ToString());
            }

            return partitions;
        }

        private double GetTotalFreeSpace(List<string> partitions)
        {
            double totalFreeSpace = 0;

            foreach (var partition in partitions)
            {
                var searcher = new ManagementObjectSearcher($"associators of {{Win32_DiskPartition.DeviceID='{partition}'}} where AssocClass=Win32_LogicalDiskToPartition");
                foreach (var obj in searcher.Get())
                {
                    var freeSpace = Convert.ToDouble(obj["FreeSpace"]);
                    totalFreeSpace += freeSpace;
                }
            }

            return totalFreeSpace;
        }

        private string GetMemoryType(uint speed)
        {
            if (speed >= 4800 && speed <= 8400)
                return "DDR5";
            else if (speed >= 2133 && speed <= 3200)
                return "DDR4";
            else if (speed >= 1066 && speed <= 1866)
                return "DDR3";
            else
                return "Unknown";
        }

        private string GetAnyDeskId()
        {
            try
            {
                Log("Getting AnyDesk ID...");
                var process = new Process
                {
                    StartInfo = new ProcessStartInfo
                    {
                        FileName = @"C:\Program Files (x86)\AnyDesk\AnyDesk.exe",
                        Arguments = "--get-id",
                        RedirectStandardOutput = true,
                        UseShellExecute = false,
                        CreateNoWindow = true
                    }
                };

                process.Start();
                string output = process.StandardOutput.ReadToEnd();
                process.WaitForExit();
                return output.Trim();
            }
            catch (Exception ex)
            {
                Log($"Error getting AnyDesk ID: {ex.Message}");
                return "N/A";
            }
        }

        private string GetTeamViewerId()
        {
            try
            {
                Log("Getting TeamViewer ID...");
                var regKey = Microsoft.Win32.Registry.LocalMachine.OpenSubKey(@"SOFTWARE\WOW6432Node\TeamViewer");
                if (regKey != null)
                {
                    var id = regKey.GetValue("ClientID");
                    return id != null ? id.ToString() : "N/A";
                }
                return "N/A";
            }
            catch (Exception ex)
            {
                Log($"Error getting TeamViewer ID: {ex.Message}");
                return "N/A";
            }
        }

        private dynamic GetDefaultPrinter()
        {
            try
            {
                Log("Getting default printer...");
                var process = new Process
                {
                    StartInfo = new ProcessStartInfo
                    {
                        FileName = "wmic",
                        Arguments = "printer get name,default",
                        RedirectStandardOutput = true,
                        UseShellExecute = false,
                        CreateNoWindow = true
                    }
                };

                process.Start();
                var output = process.StandardOutput.ReadToEnd();
                process.WaitForExit();

                Log($"WMIC Output: {output}");

                var lines = output.Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries);

                foreach (var line in lines)
                {
                    if (line.Contains("TRUE"))
                    {
                        var parts = line.Trim().Split(new[] { ' ', '\t' }, StringSplitOptions.RemoveEmptyEntries);
                        if (parts.Length > 1)
                        {
                            var printerName = string.Join(" ", parts, 1, parts.Length - 1);
                            Log($"Found default printer: {printerName}");
                            return GetPrinterDetails(printerName);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Log($"Error getting default printer: {ex.Message}");
            }

            return null;
        }

        private dynamic GetPrinterDetails(string printerName)
        {
            try
            {
                var searcher = new ManagementObjectSearcher($"SELECT * FROM Win32_Printer WHERE Name = '{printerName}'");

                foreach (var obj in searcher.Get())
                {
                    var portName = obj["PortName"].ToString();
                    var ipAddress = GetPrinterIpAddress(portName);

                    return new
                    {
                        Name = obj["Name"].ToString(),
                        PortName = portName,
                        Default = (bool)obj["Default"],
                        IpAddress = ipAddress
                    };
                }
            }
            catch (Exception ex)
            {
                Log($"Error getting printer details for {printerName}: {ex.Message}");
            }

            return null;
        }

        private string GetPrinterIpAddress(string portName)
        {
            try
            {
                var searcher = new ManagementObjectSearcher($"SELECT * FROM Win32_TCPIPPrinterPort WHERE Name = '{portName}'");

                foreach (var obj in searcher.Get())
                {
                    return obj["HostAddress"].ToString();
                }
            }
            catch (Exception ex)
            {
                Log($"Error getting printer IP address for port {portName}: {ex.Message}");
            }

            return "N/A";
        }

        private void Log(string message)
        {
            try
            {
                var logPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "service_log.txt");

                // Убедитесь, что директория существует
                var logDirectory = Path.GetDirectoryName(logPath);
                if (!Directory.Exists(logDirectory))
                {
                    Directory.CreateDirectory(logDirectory);
                }

                using (var writer = new StreamWriter(logPath, true))
                {
                    writer.WriteLine($"{DateTime.Now}: {message}");
                    writer.Flush();
                }
            }
            catch (Exception ex)
            {
                try
                {
                    EventLog.WriteEntry("ComputerMonitorService", $"Failed to log to file: {ex.Message}", EventLogEntryType.Error);
                }
                catch
                {
                    // Игнорируем ошибки при записи в журнал событий
                }
            }
        }
    }
}
