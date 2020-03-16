using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Configuration.Install;
using Microsoft.Win32;
namespace YanBo
{
    [RunInstaller(true)]
    public partial class Installer1 : Installer
    {
        public Installer1()
        {
            InitializeComponent();
        }
        public override void Install(IDictionary stateSaver)
        {
            base.Install(stateSaver);
            stateSaver.Add("TargetDir", Context.Parameters["DP_TargetDir"].ToString());
        }
        public override void Commit(IDictionary savedState)
        {
            base.Commit(savedState);
            String installdir = savedState["TargetDir"].ToString();
            String installPath = installdir + "YanBo.dll";
            AddToReg("YanBo", "≥Ã–Ú√Ë ˆ", installPath);
        }
        public override void Uninstall(IDictionary savedState)
        {
            DelFromReg("YanBo");
            base.Uninstall(savedState);
        }
        public static void AddToReg(String dname, String desc, String dpath)
        {
            try
            {
                RegistryKey LocalMachine = Registry.LocalMachine;
                RegistryKey Applications = LocalMachine.OpenSubKey(@"SOFTWARE\Autodesk\AutoCAD\R18.0\ACAD-8001:804\Applications", true);
                RegistryKey MyPrograrm = Applications.CreateSubKey(dname);
                MyPrograrm.SetValue("DESCRIPTION", desc, RegistryValueKind.String);
                MyPrograrm.SetValue("LOADCTRLS", 14, RegistryValueKind.DWord);
                MyPrograrm.SetValue("LOADER", dpath, RegistryValueKind.String);
                MyPrograrm.SetValue("MANAGED", 1, RegistryValueKind.DWord);

            }
            catch
            {

            }
        }
        private static void DelFromReg(String dname)
        {
            RegistryKey rk = Registry.LocalMachine;
            RegistryKey rk0 = rk.OpenSubKey("SOFTWARE\\Autodesk\\AutoCAD\\R18.0\\ACAD-8001:804\\Applications", true);
            String[] subkeys = rk0.GetSubKeyNames();
            List<String> keys = new List<String>();
            keys.AddRange(subkeys);
            if (keys.Contains(dname))
            {
                rk0.DeleteSubKeyTree(dname);
            }
        }
    }
}