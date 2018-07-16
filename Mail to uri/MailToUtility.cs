using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;

namespace Mail_to_uri
{
    public class MailToUtility
    {
        public static string SendMailToURI(List<string> vAddress, List<string> vCC, List<string> vBCC, string strSubject, string strBody)
        {
            string strURL = @"mailto:";

            // address
            if (vAddress != null && vAddress.Count > 0)
            {
                string strAddress = string.Join(";", vAddress);
                if (!string.IsNullOrEmpty(strAddress))
                    strURL += Uri.EscapeUriString(strAddress);
            }
            strURL += @"?";

            /// Encoding see:
            /// https://stackoverflow.com/a/1517709/3576052

            // subject
            if (!string.IsNullOrEmpty(strSubject))
                strURL += @"subject=" + Uri.EscapeDataString(strSubject);

            // body
            if (!string.IsNullOrEmpty(strBody))
                strURL += @"&body=" + Uri.EscapeDataString(strBody);

            // CC
            if (vCC != null && vCC.Count > 0)
            {
                string strCC = string.Join(";", vCC);
                if (!string.IsNullOrEmpty(strCC))
                    strURL += @"&cc=" + Uri.EscapeUriString(strCC);
            }

            // BCC        
            if (vBCC != null && vBCC.Count > 0)
            {
                string strBCC = string.Join(";", vBCC);
                if (!string.IsNullOrEmpty(strBCC))
                    strURL += @"&bcc=" + Uri.EscapeUriString(strBCC);
            }

            try
            {
                System.Diagnostics.Process.Start(strURL);
            }
            catch (Exception ex)
            {
                return ex.ToString();
            }
            return string.Empty;
        }

        public static bool IsMailClientInstalledRoot()
        {
            bool bReturn = false;
            bool bReturn2 = false;
            string subKeyPath = @"mailto\shell\";
            try
            {
                RegistryKey key = Registry.ClassesRoot.OpenSubKey(subKeyPath);
                if (key != null)
                {
                    object objValue = key.GetValue("");
                    if (objValue != null)
                    {
                        if (key.GetValueKind("") == RegistryValueKind.String)
                        {
                            string strValue = Convert.ToString(objValue);
                            if (strValue != @"(value not set)")
                                bReturn = true;
                        }
                    }
                    key.Close();
                    key.Dispose();
                    key = null;
                    objValue = null;

                    // 2nd test
                    subKeyPath = @"mailto\shell\open\command\";
                    key = Registry.ClassesRoot.OpenSubKey(subKeyPath);
                    if (key != null)
                    {
                        objValue = key.GetValue("");
                        if (objValue != null)
                        {
                            if (key.GetValueKind("") == RegistryValueKind.String)
                            {
                                string strValue = Convert.ToString(objValue);
                                if (strValue != null)
                                {
                                    if (strValue.IndexOf(@"C:\windows\system", StringComparison.OrdinalIgnoreCase) >= 0)
                                        bReturn2 = false;
                                    else
                                        bReturn2 = true;
                                }
                            }
                        }
                        key.Close();
                        key.Dispose();
                        key = null;
                    }
                }
            }
            catch { }
            return (bReturn || bReturn2);
        }

        public static bool IsMailClientInstalledForUser()
        {
            bool bReturn = false;
            bool bReturn2 = false;
            string subKeyPath = @"SOFTWARE\Classes\mailto\shell\";
            try
            {
                RegistryKey key = Registry.LocalMachine.OpenSubKey(subKeyPath);
                if (key != null)
                {
                    object objValue = key.GetValue("");
                    if (objValue != null)
                    {
                        if (key.GetValueKind("") == RegistryValueKind.String)
                        {
                            string strValue = Convert.ToString(objValue);
                            if (strValue != @"(value not set)")
                                bReturn = true;
                        }
                    }
                    key.Close();
                    key.Dispose();
                    key = null;
                    objValue = null;

                    // 2nd test
                    subKeyPath = @"SOFTWARE\Classes\mailto\shell\open\command\";
                    key = Registry.LocalMachine.OpenSubKey(subKeyPath);
                    if (key != null)
                    {
                        objValue = key.GetValue("");
                        if (objValue != null)
                        {
                            if (key.GetValueKind("") == RegistryValueKind.String)
                            {
                                string strValue = Convert.ToString(objValue);
                                if (strValue != null)
                                {
                                    if (strValue.IndexOf(@"C:\windows\system", StringComparison.OrdinalIgnoreCase) >= 0)
                                        bReturn2 = false;
                                    else
                                        bReturn2 = true;
                                }
                            }
                        }
                        key.Close();
                        key.Dispose();
                        key = null;
                    }
                }
            }
            catch { }
            return (bReturn || bReturn2);
        }

        /// <summary>
        /// Test if HKEY_CLASSES_ROOT\mailto\shell\ has value not set to (value not set)
        /// </summary>
        /// <returns></returns>
        public static bool TestShellValue()
        {
            // HKEY_CLASSES_ROOT\mailto\shell\
            bool bReturn = false;
            string subKeyPath = @"mailto\shell\";
            try
            {
                RegistryKey key = Registry.ClassesRoot.OpenSubKey(subKeyPath);
                if (key != null)
                {
                    object objValue = key.GetValue("");
                    if (objValue != null)
                    {
                        if (key.GetValueKind("") != RegistryValueKind.String)
                        {
                            bReturn = false;
                            MessageBox.Show("Default mailto value is not a string type");
                        }
                        else
                        {
                            string strValue = Convert.ToString(objValue);
                            if (strValue == @"(value not set)")
                            {
                                bReturn = false;
                            }
                            else
                            {
                                bReturn = true;                                
                            }
                            MessageBox.Show("Default mailto value is " + strValue);
                        }
                    }
                    else
                    {
                        MessageBox.Show("Default value does not exist");
                        bReturn = false;
                    }
                    key.Close();
                    key.Dispose();
                    key = null;
                }
                else
                    MessageBox.Show("HKEY_CLASSES_ROOT\\mailto\\shell\\\ndoes not exist");
            }
            catch { }
            return bReturn;
        }

        /// <summary>
        /// Test if HKEY_CLASSES_ROOT\mailto\shell\open\command\ has value not set to C:\windows\system32\
        /// </summary>
        /// <returns></returns>
        public static bool TestOpenCommand()
        {
            bool bReturn = false;
            string subKeyPath = @"mailto\shell\open\command\";
            try
            {
                RegistryKey key = Registry.ClassesRoot.OpenSubKey(subKeyPath);
                if (key != null)
                {
                    object objValue = key.GetValue("");
                    if (objValue != null)
                    {
                        if (key.GetValueKind("") != RegistryValueKind.String)
                        {
                            bReturn = false;
                            MessageBox.Show("Default mailto command value is not a string type");
                        }
                        else
                        {
                            string strValue = Convert.ToString(objValue);
                            if (strValue != null)
                            {
                                if (strValue.IndexOf(@"C:\windows\system", StringComparison.OrdinalIgnoreCase) >= 0)
                                    bReturn = false;
                                else
                                    bReturn = true;
                            }
                            MessageBox.Show("Default mailto command value is " + strValue);
                        }
                    }
                    else
                    {
                        MessageBox.Show("Default command value does not exist");
                        bReturn = false;
                    }
                    key.Close();
                    key.Dispose();
                    key = null;
                }
                else
                    MessageBox.Show("HKEY_CLASSES_ROOT\\mailto\\shell\\open\\command\\\ndoes not exist");
            }
            catch { }
            return bReturn;
        }
    }
}
