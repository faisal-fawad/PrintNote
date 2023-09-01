using System;
using System.Text;
using System.Runtime.InteropServices;
using System.Security;
using System.ComponentModel;
using System.Drawing.Printing;

#pragma warning disable CS3001 // Type is not CLS-compliant

namespace CreatePrintForm
{
    /// <summary>
    /// Summary description for CreatePrintForm.
    /// </summary>
    public class CreatePrintForm
    {
        // Make a static class
        private CreatePrintForm()
        {
        }
        [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Auto)]
        internal struct structPrinterDefaults
        {
            [MarshalAs(UnmanagedType.LPTStr)] public String pDatatype;
            public IntPtr pDevMode;
            [MarshalAs(UnmanagedType.I4)] public int DesiredAccess;
        };

        [DllImport("winspool.Drv", EntryPoint = "OpenPrinter", SetLastError = true,
            CharSet = CharSet.Unicode, ExactSpelling = false, CallingConvention = CallingConvention.StdCall),
            SuppressUnmanagedCodeSecurityAttribute()]
        internal static extern bool OpenPrinter([MarshalAs(UnmanagedType.LPTStr)] string printerName, out IntPtr phPrinter, ref structPrinterDefaults pd);

        [DllImport("winspool.Drv", EntryPoint = "ClosePrinter", SetLastError = true,
           CharSet = CharSet.Unicode, ExactSpelling = false,
           CallingConvention = CallingConvention.StdCall), SuppressUnmanagedCodeSecurityAttribute()]
        internal static extern bool ClosePrinter(IntPtr phPrinter);

        // Struct for FormInfo1
        [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Auto)]
        internal struct structSize
        {
            public Int32 width;
            public Int32 height;
        }

        // Struct for FormInfo1
        [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Auto)]
        internal struct structRect
        {
            public Int32 left;
            public Int32 top;
            public Int32 right;
            public Int32 bottom;
        }

        // FormInfo1 struct
        [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
        internal struct FormInfo1
        {
            public uint Flags;
            public String pName;
            public structSize Size;
            public structRect ImageableArea;
        };

        [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Ansi/* changed from CharSet=CharSet.Auto */)]
        internal struct structDevMode
        {
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 32)]
            public String
              dmDeviceName;
            [MarshalAs(UnmanagedType.U2)] public short dmSpecVersion;
            [MarshalAs(UnmanagedType.U2)] public short dmDriverVersion;
            [MarshalAs(UnmanagedType.U2)] public short dmSize;
            [MarshalAs(UnmanagedType.U2)] public short dmDriverExtra;
            [MarshalAs(UnmanagedType.U4)] public int dmFields;
            [MarshalAs(UnmanagedType.I2)] public short dmOrientation;
            [MarshalAs(UnmanagedType.I2)] public short dmPaperSize;
            [MarshalAs(UnmanagedType.I2)] public short dmPaperLength;
            [MarshalAs(UnmanagedType.I2)] public short dmPaperWidth;
            [MarshalAs(UnmanagedType.I2)] public short dmScale;
            [MarshalAs(UnmanagedType.I2)] public short dmCopies;
            [MarshalAs(UnmanagedType.I2)] public short dmDefaultSource;
            [MarshalAs(UnmanagedType.I2)] public short dmPrintQuality;
            [MarshalAs(UnmanagedType.I2)] public short dmColor;
            [MarshalAs(UnmanagedType.I2)] public short dmDuplex;
            [MarshalAs(UnmanagedType.I2)] public short dmYResolution;
            [MarshalAs(UnmanagedType.I2)] public short dmTTOption;
            [MarshalAs(UnmanagedType.I2)] public short dmCollate;
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 32)] public String dmFormName;
            [MarshalAs(UnmanagedType.U2)] public short dmLogPixels;
            [MarshalAs(UnmanagedType.U4)] public int dmBitsPerPel;
            [MarshalAs(UnmanagedType.U4)] public int dmPelsWidth;
            [MarshalAs(UnmanagedType.U4)] public int dmPelsHeight;
            [MarshalAs(UnmanagedType.U4)] public int dmNup;
            [MarshalAs(UnmanagedType.U4)] public int dmDisplayFrequency;
            [MarshalAs(UnmanagedType.U4)] public int dmICMMethod;
            [MarshalAs(UnmanagedType.U4)] public int dmICMIntent;
            [MarshalAs(UnmanagedType.U4)] public int dmMediaType;
            [MarshalAs(UnmanagedType.U4)] public int dmDitherType;
            [MarshalAs(UnmanagedType.U4)] public int dmReserved1;
            [MarshalAs(UnmanagedType.U4)] public int dmReserved2;
        }

        [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Auto)]
        internal struct PRINTER_INFO_9
        {
            public IntPtr pDevMode;
        }

        [DllImport("winspool.Drv", EntryPoint = "AddFormW", SetLastError = true,
           CharSet = CharSet.Unicode, ExactSpelling = true,
           CallingConvention = CallingConvention.StdCall), SuppressUnmanagedCodeSecurityAttribute()]
        internal static extern bool AddForm(IntPtr phPrinter, [MarshalAs(UnmanagedType.I4)] int level, ref FormInfo1 form);

        [DllImport("winspool.Drv", EntryPoint = "DeleteForm", SetLastError = true,
            CharSet = CharSet.Unicode, ExactSpelling = false, CallingConvention = CallingConvention.StdCall),
            SuppressUnmanagedCodeSecurityAttribute()]
        internal static extern bool DeleteForm(IntPtr phPrinter, [MarshalAs(UnmanagedType.LPTStr)] string pName);

        [DllImport("kernel32.dll", EntryPoint = "GetLastError", SetLastError = false,
            ExactSpelling = true, CallingConvention = CallingConvention.StdCall),
            SuppressUnmanagedCodeSecurityAttribute()]
        internal static extern Int32 GetLastError();

        [DllImport("GDI32.dll", EntryPoint = "CreateDC", SetLastError = true,
            CharSet = CharSet.Unicode, ExactSpelling = false,
            CallingConvention = CallingConvention.StdCall),
            SuppressUnmanagedCodeSecurityAttribute()]
        internal static extern IntPtr CreateDC([MarshalAs(UnmanagedType.LPTStr)] string pDrive, [MarshalAs(UnmanagedType.LPTStr)] string pName, [MarshalAs(UnmanagedType.LPTStr)] string pOutput, ref structDevMode pDevMode);

        [DllImport("GDI32.dll", EntryPoint = "ResetDC", SetLastError = true,
            CharSet = CharSet.Unicode, ExactSpelling = false,
            CallingConvention = CallingConvention.StdCall),
            SuppressUnmanagedCodeSecurityAttribute()]
        internal static extern IntPtr ResetDC(IntPtr hDC, ref structDevMode pDevMode);

        [DllImport("GDI32.dll", EntryPoint = "DeleteDC", SetLastError = true,
            CharSet = CharSet.Unicode, ExactSpelling = false,
            CallingConvention = CallingConvention.StdCall),
        SuppressUnmanagedCodeSecurityAttribute()]
        internal static extern bool DeleteDC(IntPtr hDC);

        [DllImport("winspool.Drv", EntryPoint = "SetPrinterA", SetLastError = true,
            CharSet = CharSet.Auto, ExactSpelling = true,
            CallingConvention = CallingConvention.StdCall), SuppressUnmanagedCodeSecurityAttribute()]
        internal static extern bool SetPrinter(IntPtr hPrinter, [MarshalAs(UnmanagedType.I4)] int level, IntPtr pPrinter, [MarshalAs(UnmanagedType.I4)] int command);

        [DllImport("winspool.Drv", EntryPoint = "DocumentPropertiesA", SetLastError = true,
            ExactSpelling = true, CallingConvention = CallingConvention.StdCall)]
        public static extern int DocumentProperties(
            IntPtr hwnd /* Handle to parent window */,
            IntPtr hPrinter /* Handle to printer object */,
            [MarshalAs(UnmanagedType.LPStr)] string pDeviceName /* Device name */,
            IntPtr pDevModeOutput /* Modified device mode */,
            IntPtr pDevModeInput /* Original device mode */,
            int fMode /* Mode options */
        );

        [DllImport("winspool.Drv", EntryPoint = "GetPrinterA", SetLastError = true,
            ExactSpelling = true, CallingConvention = CallingConvention.StdCall)]
        public static extern bool GetPrinter(IntPtr hPrinter, int dwLevel, IntPtr pPrinter, int dwBuf, out int dwNeeded);

        // SendMessageTimeout tools
        [Flags]
        public enum SendMessageTimeoutFlags : uint
        {
            SMTO_NORMAL = 0x0000,
            SMTO_BLOCK = 0x0001,
            SMTO_ABORTIFHUNG = 0x0002,
            SMTO_NOTIMEOUTIFNOTHUNG = 0x0008
        }
        const int WM_SETTINGCHANGE = 0x001A;
        const int HWND_BROADCAST = 0xffff;

        [DllImport("user32.dll", SetLastError = true, CharSet = CharSet.Auto)]
        public static extern IntPtr SendMessageTimeout(
           IntPtr windowHandle,
           uint Msg,
           IntPtr wParam,
           IntPtr lParam,
           SendMessageTimeoutFlags flags,
           uint timeout,
           out IntPtr result
        );

        /// <summary>
        /// Add the printer form to a printer 
        /// </summary>
        /// <param name="printerName">The printer name</param>
        /// <param name="paperName">Name of the printer form</param>
        /// <param name="widthMm">Width given in millimeters</param>
        /// <param name="heightMm">Height given in millimeters</param>
        public static void AddCustomPaperSize(string printerName, string paperName, float widthMm, float heightMm)
        {
            // The code to add a custom paper size is different for Windows NT then it is for previous versions of windows
            if (PlatformID.Win32NT == Environment.OSVersion.Platform) 
            {
                const int PRINTER_ACCESS_USE = 0x00000008;
                const int PRINTER_ACCESS_ADMINISTER = 0x00000004;

                structPrinterDefaults defaults = new structPrinterDefaults();
                defaults.pDatatype = null;
                defaults.pDevMode = IntPtr.Zero;
                defaults.DesiredAccess = PRINTER_ACCESS_ADMINISTER | PRINTER_ACCESS_USE;
                IntPtr hPrinter;

                // Open the printer
                if (OpenPrinter(printerName, out hPrinter, ref defaults))
                {
                    try
                    {
                        // Delete form incase it already exists
                        DeleteForm(hPrinter, paperName);

                        // Create and initialize the FORM_INFO_1 structure
                        FormInfo1 formInfo = new FormInfo1();
                        formInfo.Flags = 0;
                        formInfo.pName = paperName;
                        // All sizes are 1000th of a millimeter
                        formInfo.Size.width = (int)(widthMm * 1000.0);
                        formInfo.Size.height = (int)(heightMm * 1000.0);
                        formInfo.ImageableArea.left = 0;
                        formInfo.ImageableArea.right = formInfo.Size.width;
                        formInfo.ImageableArea.top = 0;
                        formInfo.ImageableArea.bottom = formInfo.Size.height;
                        if (!AddForm(hPrinter, 1, ref formInfo))
                        {
                            StringBuilder strBuilder = new StringBuilder();
                            strBuilder.AppendFormat("Failed to add the custom paper size {0} to the printer {1}, System error number: {2}", paperName, printerName, GetLastError());
                            throw new ApplicationException(strBuilder.ToString());
                        }

                        // Initialization
                        const int DM_OUT_BUFFER = 2;
                        const int DM_IN_BUFFER = 8;
                        structDevMode devMode = new structDevMode();
                        IntPtr hPrinterInfo, hDummy;
                        PRINTER_INFO_9 printerInfo;
                        printerInfo.pDevMode = IntPtr.Zero;
                        int iPrinterInfoSize, iDummyInt;

                        // Get the size of the DEV_MODE buffer
                        int iDevModeSize = DocumentProperties(IntPtr.Zero, hPrinter, printerName, IntPtr.Zero, IntPtr.Zero, 0);
                        if (iDevModeSize < 0)
                            throw new ApplicationException("Cannot get the size of the DEVMODE structure.");

                        // Allocate the DEV_MODE buffer
                        IntPtr hDevMode = Marshal.AllocCoTaskMem(iDevModeSize + 100);

                        // Get a pointer to the DEV_MODE buffer
                        int iRet = DocumentProperties(IntPtr.Zero, hPrinter, printerName, hDevMode, IntPtr.Zero, DM_OUT_BUFFER);
                        if (iRet < 0)
                            throw new ApplicationException("Cannot get the DEVMODE structure.");

                        // Fill the DEV_MODE structure
                        devMode = (structDevMode)Marshal.PtrToStructure(hDevMode, devMode.GetType());

                        // Set the form name fields which indicates that the field will be modified
                        devMode.dmFields = 0x10000;
                        devMode.dmFormName = paperName;

                        // Put the DEV_MODE structure back into the pointer and merge the new changes with the old
                        Marshal.StructureToPtr(devMode, hDevMode, true);
                        iRet = DocumentProperties(IntPtr.Zero, hPrinter, printerName, printerInfo.pDevMode, printerInfo.pDevMode, DM_IN_BUFFER | DM_OUT_BUFFER);
                        if (iRet < 0)
                            throw new ApplicationException("Unable to edit the orientation setting for this printer");

                        // Get the size of PRINTER_INFO_9 structure
                        GetPrinter(hPrinter, 9, IntPtr.Zero, 0, out iPrinterInfoSize);
                        if (iPrinterInfoSize == 0)
                            throw new ApplicationException("GetPrinter failed. Couldn't get the size of the PRINTER_INFO_9 structure");

                        // Allocate the PRINTER_INFO_9 buffer 
                        hPrinterInfo = Marshal.AllocCoTaskMem(iPrinterInfoSize + 100);

                        // Get a pointer to the PRINTER_INFO_9 buffer
                        bool bSuccess = GetPrinter(hPrinter, 9, hPrinterInfo, iPrinterInfoSize, out iDummyInt);
                        if (!bSuccess)
                            throw new ApplicationException("GetPrinter failed. Couldn't get the PRINTER_INFO_9 structure");

                        // Fill the PRINTER_INFO_9 structure
                        printerInfo = (PRINTER_INFO_9)Marshal.PtrToStructure(hPrinterInfo, printerInfo.GetType());
                        printerInfo.pDevMode = hDevMode;

                        // Get a pointer to the PRINTER_INFO_9 structure
                        Marshal.StructureToPtr(printerInfo, hPrinterInfo, true);

                        // Set the printer settings using the the PRINTER_INFO_9 structure
                        bSuccess = SetPrinter(hPrinter, 9, hPrinterInfo, 0);
                        if (!bSuccess)
                            throw new Win32Exception(Marshal.GetLastWin32Error(), "SetPrinter failed. Couldn't set the printer settings");

                        // Tell all open programs that this change occurred.
                        SendMessageTimeout(
                           new IntPtr(HWND_BROADCAST),
                           WM_SETTINGCHANGE,
                           IntPtr.Zero,
                           IntPtr.Zero,
                           CreatePrintForm.SendMessageTimeoutFlags.SMTO_NORMAL,
                           1000,
                           out hDummy
                        );
                    }
                    finally
                    {
                        ClosePrinter(hPrinter);
                    }
                }
                else
                {
                    StringBuilder strBuilder = new StringBuilder();
                    // System error 5 represents missing of administrative permissions 
                    strBuilder.AppendFormat("Failed to open the {0} printer, System error number: {1}", printerName, GetLastError());
                    throw new ApplicationException(strBuilder.ToString());
                }
            }
            else
            {
                structDevMode pDevMode = new structDevMode();
                IntPtr hDC = CreateDC(null, printerName, null, ref pDevMode);
                if (hDC != IntPtr.Zero)
                {
                    const long DM_PAPERSIZE = 0x00000002L;
                    const long DM_PAPERLENGTH = 0x00000004L;
                    const long DM_PAPERWIDTH = 0x00000008L;
                    pDevMode.dmFields = (int)(DM_PAPERSIZE | DM_PAPERWIDTH | DM_PAPERLENGTH);
                    pDevMode.dmPaperSize = 256;
                    pDevMode.dmPaperWidth = (short)(widthMm * 1000.0);
                    pDevMode.dmPaperLength = (short)(heightMm * 1000.0);
                    ResetDC(hDC, ref pDevMode);
                    DeleteDC(hDC);
                }
            }
        }
    }
}
