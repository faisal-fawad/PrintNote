// To add support for custom size manually follow the instructions from the following link:
// https://franklinheath.co.uk/2015/08/29/custom-page-sizes-for-microsoft-print-to-pdf/
using Microsoft.Win32;
using System.IO.Compression;

string match = "}\r\n*Feature: Orientation";
string insert = "*Option: CUSTOMSIZE\r\n{\r\n*rcNameID: =USER_DEFINED_SIZE_DISPLAY\r\n*MinSize: PAIR(180000, 180000)\r\n*MaxSize: PAIR(30276000 , 42804000)\r\n*MaxPrintableWidth: 30276000\r\n}\n";
string exist = "*Option: CUSTOMSIZE";
string path = "HKEY_LOCAL_MACHINE\\SOFTWARE\\Microsoft\\Windows NT\\CurrentVersion\\Print\\Printers\\Microsoft Print to PDF";

if (OperatingSystem.IsWindows())
{
    // Attempting to get the .gpd file location
    var tempDir = Registry.GetValue(path, "PrintQueueV4DriverDirectory", null);
    var tempGpd = Registry.GetValue(path + "\\PrinterDriverData", "V4_Merged_ConfigFile_Name", null);
    if (tempGpd == null | tempDir == null)
    {
        Console.WriteLine("Unable to find .gpd file location");
        Console.Read();
        return;
    }
    string gpdFile = Environment.GetEnvironmentVariable("windir") + "\\System32\\spool\\V4Dirs\\" + tempDir + "\\" + tempGpd;
    string gpdDir = Environment.GetEnvironmentVariable("windir") + "\\System32\\spool\\V4Dirs\\" + tempDir;
    Console.WriteLine("Using .gpd file location: " + gpdFile);
 
    try
    {
        // Checks if the .gpd file already has support for custom sizes
        string gpdContents = File.ReadAllText(gpdFile);
        if (gpdContents.Contains(exist))
        {
            Console.WriteLine("Microsoft Print to PDF already supports custom sizes!");
            Console.Read();
            return;
        }

        // Saves a backup of the folder before editing the .gpd file
        if (File.Exists(gpdDir + ".zip"))
        {
            File.Delete(gpdDir + ".zip");
        }
        ZipFile.CreateFromDirectory(gpdDir, gpdDir + ".zip");

        // Adds support for custom sizes to the .gpd file
        int index = gpdContents.IndexOf(match);
        string newGpd = gpdContents.Insert(index, insert);
        File.WriteAllText(gpdFile, newGpd);
        Console.WriteLine("Success!", ConsoleColor.Green);
    }
    catch (Exception e)
    {
        Console.WriteLine("Unknown exception: " + e.Message);
    }
}
else
{
    Console.WriteLine("This program only works on Windows!");
}
Console.Read(); // Keeps the console open