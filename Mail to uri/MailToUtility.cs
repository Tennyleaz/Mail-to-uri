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
                    strURL += strAddress;
            }
            strURL += @"?";

            // subject
            if (!string.IsNullOrEmpty(strSubject))
                strURL += @"subject=" + System.Web.HttpUtility.UrlEncode(strSubject);

            // body
            if (!string.IsNullOrEmpty(strBody))
                strURL += @"&body=" + System.Web.HttpUtility.UrlEncode(strBody);

            // CC
            if (vCC != null && vCC.Count > 0)
            {
                string strCC = string.Join(";", vCC);
                if (!string.IsNullOrEmpty(strCC))
                    strURL += @"&cc=" + strCC;
            }

            // BCC        
            if (vBCC != null && vBCC.Count > 0)
            {
                string strBCC = string.Join(";", vBCC);
                if (!string.IsNullOrEmpty(strBCC))
                    strURL += @"&bcc=" + strBCC;
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

        /// <summary>
        /// Test if HKEY_CLASSES_ROOT\mailto\shell\ has value not set to (value not set)
        /// </summary>
        /// <returns></returns>
        public static bool IsMailClientInstalled()
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
