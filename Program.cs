using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Win32;

//Namespace
namespace Outlook_disable_attachment_preview
{
    class Program
    {
        public static void Main(string[] args)
        {
            //Set app vars
            string outlookinstallpath = string.Empty; // String used for get Outlook installed path in registry
            string outlookver = string.Empty; // String used for showing Outlook version using installed path details in registry
            string outlookdisableattachpreviewpath = string.Empty; // String used for get Outlook attachment preview registry path
            string[] outlookinfo = new string[3];

            //Pass back return value from function GetAppInfo
            outlookinfo = GetAppInfo(outlookinstallpath, outlookver, outlookdisableattachpreviewpath);
            //Check Outlook installation error handling, and run respective functions
            if (outlookinfo[0]=="")
                {
                Console.WriteLine("Quitting program as Outlook installation is not detected.");
                }
            else 
            {
              Menu_changes(outlookinfo[0]);// Run GetAppInfo function to get info
             }
            Console.WriteLine();
            Console.WriteLine("Done! Press any key to continue....");
            string input = Console.ReadLine();
            return;
        }

        // Get and display app info
        static string[] GetAppInfo(string outlookinstallpath1, string outlookver1, string outlookdisableattachpreviewpath1)
        {
            


            // Set app vars
            string appName = "Outlook attachment preview disable tool";
            string appVersion = "1.0.0";
            string appAuthor = "Velan Ramalinggam (velanr@gmail.com)";
            string appdesc1 = "This tool helps enable / disable email attachment preview function in Microsoft Outlook. ";
            string appdesc2 = "It is useful as it can help eliminate malware execution by attachment preview of malware infected Office documents/macro via unpatched Office bugs.";
            string appcoverage = "Covers Microsoft Outlook version 2010, 2013 & 2016. \nIt DOES NOT cover Microsoft Outlook for Office 365.";
            //Console.ForegroundColor = ConsoleColor.Red;            
            string appnote = "Note : Make sure to run this program in administrator mode (Run as administrator)";

            //Get Outlook path & version
            outlookinstallpath1 = ReadSubKeyValue_LocalMachine(@"SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\OUTLOOK.EXE", "Path");

            //Get Outlook disable attachment preview registry path info
            //outlookdisableattachpreviewpath = ReadSubKeyValue_CurrentUser(@"SOFTWARE\Policies\Microsoft\office\16.0\outlook\preferences", "disableattachmentpreviewing");

            // Change text color
            Console.ForegroundColor = ConsoleColor.Green;

            //Sort Outlook version, and existence of disable attachement preview key, in user friendly manner
            if(outlookinstallpath1.Contains("Office16")) {
                outlookver1 = "2016";
                outlookdisableattachpreviewpath1 = ReadSubKeyValue_CurrentUser(@"SOFTWARE\Policies\Microsoft\office\16.0\outlook\preferences", "disableattachmentpreviewing");
                if (outlookdisableattachpreviewpath1 == null)
                {
                    outlookdisableattachpreviewpath1 = "Not exist/defined in registry";
                }
            }
                else if (outlookinstallpath1.Contains("Office15"))
            {
                outlookver1 = "2013";
                outlookdisableattachpreviewpath1 = ReadSubKeyValue_CurrentUser(@"SOFTWARE\Policies\Microsoft\office\15.0\outlook\preferences", "disableattachmentpreviewing");
                if(outlookdisableattachpreviewpath1 == null)
                {
                    outlookdisableattachpreviewpath1 = "Not exist/defined in registry";
                }
            }
            else if (outlookinstallpath1.Contains("Office14"))
            {
                outlookver1 = "2010";
                outlookdisableattachpreviewpath1 = ReadSubKeyValue_CurrentUser(@"SOFTWARE\Policies\Microsoft\office\14.0\outlook\preferences", "disableattachmentpreviewing");
                if (outlookdisableattachpreviewpath1 == null)
                {
                    outlookdisableattachpreviewpath1 = "Not exist/defined in registry";
                }
            }
            
            else if(outlookinstallpath1.Contains(""))
            {
                outlookdisableattachpreviewpath1 = "Not exist/defined in registry";
                Console.WriteLine("Outlook installation NOT DETECTED ! Are you sure Outlook is installed ?");
            }
            else
            {
            //outlookver1 = "????";
            //outlookdisableattachpreviewpath1 = "Not defined";
            }
            string outlookdisableattachpreviewpath2 = String.Empty;

            if (outlookdisableattachpreviewpath1 == "1")
            {
                outlookdisableattachpreviewpath2 = "\nEnabled (1) - Disable attachment preview is enabled, so attachment preview is disabled"; 
                    }
            else if (outlookdisableattachpreviewpath1 == "0")
            {
                outlookdisableattachpreviewpath2 = "\nDisabled (0) - Disable attachment preview is disabled, so attachment preview is enabled";
            }
            else if (outlookdisableattachpreviewpath1 == "Not exist/defined in registry")
            {
                outlookdisableattachpreviewpath2 = "Not exist/defined in registry. Default setting is enable attachement preview.";
            }
            //else
            //{
            //    outlookdisableattachpreviewpath2 = "Not exist";
            //}
            // Write out app info
            Console.WriteLine("{0}: Version {1} by {2}", appName, appVersion, appAuthor);
            Console.WriteLine();
            Console.WriteLine("{0}", appdesc1);
            Console.WriteLine("{0}", appdesc2);
            Console.WriteLine();
            Console.WriteLine("{0}", appcoverage);
            Console.WriteLine();
            Console.WriteLine("Outlook path : {0}", outlookinstallpath1);
            Console.WriteLine("Outlook version : Microsoft Outlook {0}", outlookver1);
            Console.WriteLine("Outlook disable attachement preview : {0}", outlookdisableattachpreviewpath2);
            Console.WriteLine();
            Console.WriteLine("Outlook disable attachment preview registry path : ");
            Console.WriteLine("SubKey HKCU\\SOFTWARE\\Policies\\Microsoft\\office\\<version number>\\outlook\\preferences, Key disableattachmentpreviewing");
            Console.WriteLine();
            Console.WriteLine("{0}", appnote);
            Console.WriteLine();



            // Reset text color
            Console.ResetColor();

            //Reference - https://www.dotnetperls.com/array
            string[] returnstring = new string[3];
            returnstring[0] = outlookver1;
            returnstring[1] = outlookdisableattachpreviewpath1;
            returnstring[2] = outlookinstallpath1;

            //outlookver = outlookver1;
            //outlookdisableattachpreviewpath = outlookdisableattachpreviewpath1;
            //outlookinstallpath = outlookinstallpath1;
            return returnstring;
        }

        static void Menu_changes(string outlookver2)
        {
            // Change text color
            Console.ForegroundColor = ConsoleColor.Yellow;

            Console.WriteLine("Do you want to continue enabling & disabling the Outlook attachment preview feature?");
            Console.WriteLine("Type ENABLE or DISABLE or EXIT:");
            string input1 = string.Empty;
            //input1 = Console.ReadLine().ToUpper();

            do
            {
                input1 = Console.ReadLine().ToUpper();
                if (input1.Contains("ENABLE"))
                {
                    break;
                }
                if (input1.Contains("DISABLE"))
                {
                    break;
                }
                if (input1.Contains("EXIT"))
                {
                    break;
                }

                Console.WriteLine("\nInvalid input. Please try again.");
                Console.WriteLine("Type ENABLE or DISABLE or EXIT:");
                //Usage of contain may not be accurate, as any containing EXIT like BREXIT is accepted

            } while (true);

                //string input1 = string.Empty;
                //input1="ENABLE";

                if (input1 == "DISABLE" && outlookver2 == "2016")
            {
                RegistryKey key = Registry.CurrentUser.CreateSubKey("Software\\Policies\\Microsoft\\office");
                key = Registry.CurrentUser.CreateSubKey("Software\\Policies\\Microsoft\\office\\16.0");
                key = Registry.CurrentUser.CreateSubKey("Software\\Policies\\Microsoft\\office\\16.0\\outlook");
                key = Registry.CurrentUser.CreateSubKey("Software\\Policies\\Microsoft\\office\\16.0\\outlook\\preferences");
                using (key)
                {

                    key.SetValue("disableattachmentpreviewing", "00000001", RegistryValueKind.DWord);
                    key.Close();
                }
                //SetSubKeyValue_CurrentUser(@"Software\Policies\Microsoft\office\16.0\outlook\preferences", "disableattachmentpreviewing", "1");
            }

            if (input1 == "DISABLE" && outlookver2 == "2013")
            {
                RegistryKey key = Registry.CurrentUser.CreateSubKey("Software\\Policies\\Microsoft\\office");
                key = Registry.CurrentUser.CreateSubKey("Software\\Policies\\Microsoft\\office\\15.0");
                key = Registry.CurrentUser.CreateSubKey("Software\\Policies\\Microsoft\\office\\15.0\\outlook");
                key = Registry.CurrentUser.CreateSubKey("Software\\Policies\\Microsoft\\office\\15.0\\outlook\\preferences");
                using (key)
                {

                    key.SetValue("disableattachmentpreviewing", "00000001", RegistryValueKind.DWord);
                    key.Close();
                }
            }

            if (input1 == "DISABLE" && outlookver2 == "2010")
            {
                RegistryKey key = Registry.CurrentUser.CreateSubKey("Software\\Policies\\Microsoft\\office");
                key = Registry.CurrentUser.CreateSubKey("Software\\Policies\\Microsoft\\office\\14.0");
                key = Registry.CurrentUser.CreateSubKey("Software\\Policies\\Microsoft\\office\\14.0\\outlook");
                key = Registry.CurrentUser.CreateSubKey("Software\\Policies\\Microsoft\\office\\14.0\\outlook\\preferences");
                using (key)
                {

                    key.SetValue("disableattachmentpreviewing", "00000001", RegistryValueKind.DWord);
                    key.Close();
                }
            }

            if (input1 == "ENABLE" && outlookver2 == "2016")
            {
                
                RegistryKey key = Registry.CurrentUser.CreateSubKey("Software\\Policies\\Microsoft\\office");
                key = Registry.CurrentUser.CreateSubKey("Software\\Policies\\Microsoft\\office\\16.0");
                key = Registry.CurrentUser.CreateSubKey("Software\\Policies\\Microsoft\\office\\16.0\\outlook");
                key = Registry.CurrentUser.CreateSubKey("Software\\Policies\\Microsoft\\office\\16.0\\outlook\\preferences");
                using (key)
                {

                    key.SetValue("disableattachmentpreviewing", "00000000", RegistryValueKind.DWord);
                    key.Close();
                }
                //SetSubKeyValue_CurrentUser(@"Software\Policies\Microsoft\office\16.0\outlook\preferences", "disableattachmentpreviewing", "1");
            }

            if (input1 == "ENABLE" && outlookver2 == "2013")
            {
                RegistryKey key = Registry.CurrentUser.CreateSubKey("Software\\Policies\\Microsoft\\office");
                key = Registry.CurrentUser.CreateSubKey("Software\\Policies\\Microsoft\\office\\15.0");
                key = Registry.CurrentUser.CreateSubKey("Software\\Policies\\Microsoft\\office\\15.0\\outlook");
                key = Registry.CurrentUser.CreateSubKey("Software\\Policies\\Microsoft\\office\\15.0\\outlook\\preferences");
                using (key)
                {

                    key.SetValue("disableattachmentpreviewing", "00000000", RegistryValueKind.DWord);
                    key.Close();
                }
            }

            if (input1 == "ENABLE" && outlookver2 == "2010")
            {
               RegistryKey key = Registry.CurrentUser.CreateSubKey("Software\\Policies\\Microsoft\\office");
                key = Registry.CurrentUser.CreateSubKey("Software\\Policies\\Microsoft\\office\\14.0");
                key = Registry.CurrentUser.CreateSubKey("Software\\Policies\\Microsoft\\office\\14.0\\outlook");
                key = Registry.CurrentUser.CreateSubKey("Software\\Policies\\Microsoft\\office\\14.0\\outlook\\preferences");
                using (key)
                {

                    key.SetValue("disableattachmentpreviewing", "00000000", RegistryValueKind.DWord);
                    key.Close();
                }
            }

            else 
            {
                return;
            }


        }

        // Check registry key & subkey value in LocalMachine - https://www.infoworld.com/article/3073167/how-to-access-the-windows-registry-using-c.html
        static string ReadSubKeyValue_LocalMachine(string subKey, string key)
                {
                    string str = string.Empty;
                    using (RegistryKey registryKey = Registry.LocalMachine.OpenSubKey(subKey))
                        {
                            if (registryKey != null)
                                {
                                    str = registryKey.GetValue(key).ToString();
                                    registryKey.Close();
                                }
                        }
                 return str;
                }
        
        
        // Check registry key & subkey value in CurrentUser - https://www.infoworld.com/article/3073167/how-to-access-the-windows-registry-using-c.html
        static string ReadSubKeyValue_CurrentUser(string subKey, string key)
        {
            string str = string.Empty;
            using (RegistryKey registryKey = Registry.CurrentUser.OpenSubKey(subKey))
            {
                try
            {
                str = registryKey.GetValue(key).ToString();

            }
            catch (NullReferenceException e)
            {
                str = null;
                //registryKey.Close();
                //Console.WriteLine("{0}", e);
            }
            catch (ArgumentNullException e)
            {
                str = null;
                //registryKey.Close();
                //Console.WriteLine("{0}", e);
            }
            //using (RegistryKey registryKey = Registry.CurrentUser.OpenSubKey(subKey))
            //{
            //    if (registryKey != null)
            //    {
            //        str = registryKey.GetValue(key).ToString();
            //        registryKey.Close();
            //    }
            //    else if (registryKey == null)
            //    {
            //        str = "None";
            //        return str;
            //
            //    }
                
            
            }
            
            return str;
        }

        //private static void ShowErrorMessage(Exception e, string v)
        //{
        //    throw new NotImplementedException();
        //}

        // Set registry key & subkey value in CurrentUser - https://www.infoworld.com/article/3073167/how-to-access-the-windows-registry-using-c.html
        static void SetSubKeyValue_CurrentUser(string subKey, string key, string keyvalue)
        {
            
            
            //string str = string.Empty;
            using (RegistryKey registryKey = Registry.CurrentUser.OpenSubKey(subKey))
            {
                //if (registryKey != null)
                //{
                    registryKey.SetValue(key, keyvalue, RegistryValueKind.DWord);
                    registryKey.Close();
                //}
                //if (registryKey == null)
                //{
                //    registryKey.SetValue(key, keyvalue);
                //    registryKey.Close();
                //}
            }
            //return str;
        }


    }
}
