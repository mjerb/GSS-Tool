using System;
using Microsoft.Win32;

namespace GSS_Alert_Service
{
    static class RegistryEngine
    {
        private static RegistryKey key = Registry.LocalMachine.OpenSubKey("Software\\GSSTools");

        public static string ReadRegistry(string KeyName)
        {
            Object value = key.GetValue(KeyName);
            if (value != null)
            {
                return value as string;
            }
            else
            {
                throw new Exception();
            }
        }

        public static void WriteRegistry(string KeyName, string KeyValue)
        {
            key.CreateSubKey(KeyName);
            RegistryKey edit = key.OpenSubKey(KeyName, true);
            edit.SetValue(KeyName, KeyValue);
        }
    }
}
