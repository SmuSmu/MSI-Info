using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using WindowsInstaller;

//Console.WriteLine("Hello World!");

namespace MSI_Info
{
    class Program
    {
        struct MsiInfo
        {
            static void Main(string[] args)
            {
                //Refactor this later to be cleaner.

                Type type = Type.GetTypeFromProgID("WindowsInstaller.Installer");

                WindowsInstaller.Installer installer = (WindowsInstaller.Installer)

                Activator.CreateInstance(type);

                WindowsInstaller.Database db;

                try
                {
                    db = installer.OpenDatabase("..\\..\\..\\..\\TestFile\\IPFilter.msi", 0);
                }
                catch (Exception e)
                {
                    throw new Exception("File does not exist or is not accessible", e);
                }

                Console.WriteLine("[Property]");
                WindowsInstaller.View dv = db.OpenView("SELECT `Property`, `Value` FROM `Property`");

                dv.Execute();

                for (WindowsInstaller.Record row; (row = dv.Fetch()) != null; )
                {
                    Console.WriteLine(row.get_StringData(1).ToString() + "=" + row.get_StringData(2).ToString());
                }
            }
        }
    }
}
 