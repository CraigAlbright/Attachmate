using System;
using System.Diagnostics;
using System.Linq;
using System.Runtime.InteropServices;
using EnvDTE;

namespace ProcessAttachMate
{
    public static class AttachMate
    {
        //Vs2012
        private const string ProgramId = @"VisualStudio.DTE.11.0";
        //Vs2010
        //private const string ProgramId = @"VisualStudio.DTE.10.0";
        //Vs2008
        //private const string ProgramId = @"VisualStudio.DTE.9.0";
        //Todo make configurable
        private static bool DebugServer = true;
        private static bool _clientAttached;
        private static bool _serverAttached;
        
        public static void Attach()
        {
            // Reference Visual Studio core
            DTE dte;
            
            try
            {
                dte = (DTE)Marshal.GetActiveObject(ProgramId);
            }
            catch (COMException)
            {
                Debug.WriteLine("Could not find Visual Studio {0}",ProgramId);
                return;
            }

            // Todo make this configurable
            var tryCount = 5;
            while (tryCount-- > 0)
            {
                try
                {
                    var processes = dte.Debugger.LocalProcesses;
                    
                    if(!_clientAttached)
                    ConnectToClientSideCode(processes);
                    if (DebugServer && !_serverAttached)
                        ConnectToServerSideCode(processes);
                }
                catch (Exception)
                {
                    //MessageBox.Show(e.Message);
                    System.Threading.Thread.Sleep(1000);
                }
            }
        }
        /// <summary>
        /// Currently this is for RightAngle.exe, but 
        /// </summary>
        /// <param name="processes"></param>
        private static void ConnectToClientSideCode(Processes processes)
        {
            var exeName = System.Configuration.ConfigurationManager.AppSettings["exeName1"];
            foreach (var proc in processes.Cast<EnvDTE.Process>().Where(proc =>
                                                                        proc.Name.IndexOf(exeName,
                                                                                          StringComparison
                                                                                              .Ordinal) != -1))
            {
                proc.Attach();
                _clientAttached = true;
            }
        }

        private static void ConnectToServerSideCode(Processes processes)
        {
            var exeName = System.Configuration.ConfigurationManager.AppSettings["exeName2"];
            foreach (var proc in processes.Cast<EnvDTE.Process>().Where(proc =>
                                                                        proc.Name.IndexOf(exeName,
                                                                                          StringComparison.Ordinal) !=
                                                                        -1))
            proc.Attach();
            _serverAttached = true;
        }
    }
}
