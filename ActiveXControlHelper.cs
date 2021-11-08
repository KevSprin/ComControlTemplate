using Microsoft.Win32;
using System;
using System.Diagnostics;
using System.Reflection;
using System.Runtime.InteropServices;

namespace RegFreeActiveXControls
{

    /* 
     * Implements access to the InteropToolbox in similar fashion to the
     * vb version of the library. Access by using My.InteropToolbox.
     */

    internal static class ComRegistration
    {
        const int OLEMISC_RECOMPOSEONRESIZE = 1;
        const int OLEMISC_CANTLINKINSIDE = 16;
        const int OLEMISC_INSIDEOUT = 128;
        const int OLEMISC_ACTIVATEWHENVISIBLE = 256;
        const int OLEMISC_SETCLIENTSITEFIRST = 131072;

        public static void RegisterControl(Type t)
        {
            try
            {

                // CLSID
                string key = @"CLSID\" + t.GUID.ToString("B");

                using (RegistryKey subkey = Registry.ClassesRoot.OpenSubKey(key, true))
                {

                    // InProcServer32: https://docs.microsoft.com/en-us/windows/win32/com/inprocserver32
                    RegistryKey inprocKey = subkey.OpenSubKey("InprocServer32", true);
                    if (inprocKey != null)
                    {
                        inprocKey.SetValue(null, Environment.SystemDirectory + @"\mscoree.dll");
                        //inprocKey.SetValue("CodeBase", Assembly.GetExecutingAssembly().CodeBase);
                    }

                    //Control: https://docs.microsoft.com/en-us/windows/win32/com/control
                    using (var c = subkey.CreateSubKey("Control"))
                    { }

                    //Misc
                    using (RegistryKey miscKey = subkey.CreateSubKey("MiscStatus"))
                    {
                        const int MiscStatusValue = OLEMISC_RECOMPOSEONRESIZE +
                                                    OLEMISC_CANTLINKINSIDE + OLEMISC_INSIDEOUT +
                                                    OLEMISC_ACTIVATEWHENVISIBLE + OLEMISC_SETCLIENTSITEFIRST;

                        miscKey.SetValue("", MiscStatusValue.ToString(), RegistryValueKind.String);
                    }

                    // ToolBoxBitmap32: https://docs.microsoft.com/en-us/windows/win32/com/toolboxbitmap32
                    using (RegistryKey bitmapKey = subkey.CreateSubKey("ToolBoxBitmap32"))
                    {
                        // If you want to have different icons for each control in this assembly
                        // you can modify this section to specify a different icon each time.
                        // Each specified icon must be embedded as a win32resource in the
                        // assembly; the default one is at index 101, but you can add additional ones.
                        bitmapKey.SetValue("", Assembly.GetExecutingAssembly().Location + ", 101",
                                           RegistryValueKind.String);
                    }

                    //TypeLib
                    using (RegistryKey typeLibKey = subkey.CreateSubKey("TypeLib"))
                    {
                        Guid libId = Marshal.GetTypeLibGuidForAssembly(t.Assembly);
                        typeLibKey.SetValue("", libId.ToString("B"), RegistryValueKind.String);
                    }

                    //Version: https://docs.microsoft.com/en-us/windows/win32/com/version
                    using (RegistryKey versionKey = subkey.CreateSubKey("Version"))
                    {
                        int major, minor;
                        Marshal.GetTypeLibVersionForAssembly(t.Assembly, out major, out minor);
                        versionKey.SetValue("", String.Format("{0}.{1}", major, minor));
                    }

                }

                const string Source = "Host .NET Interop UserControl in VB6";
                const string Log = "Application";
                string sEvent = "Registration successful: key = " + key;

                if (!EventLog.SourceExists(Source))
                    EventLog.CreateEventSource(Source, Log);

                EventLog.WriteEntry(Source, sEvent, EventLogEntryType.Warning, 234);
            }
            catch (Exception ex)
            {
                LogAndRethrowException("ComRegisterFunction failed.", t, ex);
            }
        }

        public static void UnregisterControl(Type t)
        {
            try
            {
                // CLSID
                string key = @"CLSID\" + t.GUID.ToString("B");
                Registry.ClassesRoot.DeleteSubKeyTree(key);

            }
            catch (Exception ex)
            {
                LogAndRethrowException("ComUnregisterFunction failed.", t, ex);
            }

        }

        private static void LogAndRethrowException(string message, Type t, Exception ex)
        {
            try
            {
                if (null != t)
                {
                    message += Environment.NewLine + String.Format("CLR class '{0}'", t.FullName);
                }

                throw new ComRegistrationException(message, ex);
            }
            catch (Exception ex2)
            {
                const string Source = "Host .NET Interop UserControl in VB6";
                const string Log = "Application";
                string sEvent = t.GUID.ToString("B") + " registration failed: " + Environment.NewLine + ex2.Message;

                if (!EventLog.SourceExists(Source))
                    EventLog.CreateEventSource(Source, Log);

                EventLog.WriteEntry(Source, sEvent, EventLogEntryType.Warning, 234);
            }
        }
    }

    [Serializable]
    public class ComRegistrationException : Exception
    {
        public ComRegistrationException() { }
        public ComRegistrationException(string message, Exception inner)
            : base(message, inner)
        {
        }
    }
}