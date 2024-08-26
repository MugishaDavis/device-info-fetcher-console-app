using System;
using System.IO;
using System.Management;
using System.Text;

namespace DeviceInfoToCSV
{
    class Program
    {
        static void Main(string[] args)
        {
            string desktopPath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
            string csvFilePath = Path.Combine(desktopPath, "DeviceInfo.csv");

            // Create a StringBuilder to hold the CSV data
            StringBuilder csvContent = new StringBuilder();

            // Write the CSV header
            csvContent.AppendLine("Computer Name,Serial Number,UUID,Vendor,Model,CPU,Storage,Total RAM (GB)");

            // Get device information
            string computerName = Environment.MachineName;
            string serialNumber = GetWmiProperty("Win32_BIOS", "SerialNumber");
            string uuid = GetWmiProperty("Win32_ComputerSystemProduct", "UUID");
            string vendor = GetWmiProperty("Win32_ComputerSystem", "Manufacturer");
            string model = GetWmiProperty("Win32_ComputerSystem", "Model");
            string cpu = GetCpuInfo();
            string storage = GetStorageInfo();
            string ram = GetRamInfo();

            // Append the information to the CSV content
            csvContent.AppendLine($"{computerName},{serialNumber},{uuid},{vendor},{model},{cpu},{storage},{ram}");

            // Write the CSV content to a file on the desktop
            File.WriteAllText(csvFilePath, csvContent.ToString());

            Console.WriteLine($"Device information has been written to {csvFilePath}");
            Console.WriteLine("Press any key to exit...");
            Console.ReadKey();
        }

        static string GetWmiProperty(string wmiClass, string wmiProperty)
        {
            using (ManagementObjectSearcher searcher = new ManagementObjectSearcher($"SELECT {wmiProperty} FROM {wmiClass}"))
            {
                foreach (ManagementObject obj in searcher.Get())
                {
                    return obj[wmiProperty]?.ToString();
                }
            }
            return "Unknown";
        }

        static string GetCpuInfo()
        {
            StringBuilder cpuInfo = new StringBuilder();
            using (ManagementObjectSearcher searcher = new ManagementObjectSearcher("SELECT Name FROM Win32_Processor"))
            {
                foreach (ManagementObject obj in searcher.Get())
                {
                    cpuInfo.Append(obj["Name"]?.ToString() + "; ");
                }
            }
            return cpuInfo.ToString().TrimEnd(';', ' ');
        }

        static string GetStorageInfo()
        {
            StringBuilder storageInfo = new StringBuilder();
            using (ManagementObjectSearcher searcher = new ManagementObjectSearcher("SELECT Model, Size FROM Win32_DiskDrive"))
            {
                foreach (ManagementObject obj in searcher.Get())
                {
                    double sizeInGB = Convert.ToDouble(obj["Size"]) / (1024 * 1024 * 1024);
                    storageInfo.Append($"{obj["Model"]} ({sizeInGB:F2} GB); ");
                }
            }
            return storageInfo.ToString().TrimEnd(';', ' ');
        }

        static string GetRamInfo()
        {
            double totalRam = 0;
            using (ManagementObjectSearcher searcher = new ManagementObjectSearcher("SELECT Capacity FROM Win32_PhysicalMemory"))
            {
                foreach (ManagementObject obj in searcher.Get())
                {
                    totalRam += Convert.ToDouble(obj["Capacity"]);
                }
            }
            return (totalRam / (1024 * 1024 * 1024)).ToString("F2");
        }
    }
}
//end