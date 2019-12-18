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

                WindowsInstaller.Database db = null;

                try
                {
                    db = installer.OpenDatabase("..\\..\\..\\..\\TestFile\\IPFilter.msi", 0);
                }
                catch (Exception e)
                {
                    throw new Exception("File does not exist or is not accessible");
                }

                //Get Product Version
                WindowsInstaller.View dv = db.OpenView("SELECT `Value` FROM `Property` WHERE `Property`='ProductVersion'");

                WindowsInstaller.Record record = null;

                dv.Execute(record);

                record = dv.Fetch();

                //ProductVersion = record.get_StringData(1).ToString();
                Console.WriteLine(record.get_StringData(1).ToString());


                Console.WriteLine("[Property]");
                //Get Product Name
                dv = db.OpenView("SELECT `Property`, `Value` FROM `Property`");

                record = null;

                dv.Execute(record);


                record = dv.Fetch();

                //ProductName = record.get_StringData(1).ToString();

                Console.WriteLine(record.get_StringData(1).ToString() + "=" + record.get_StringData(2).ToString());
            }

        }
    }
}