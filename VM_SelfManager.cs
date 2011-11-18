using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.ServiceProcess;
using System.Management;
using System.Threading;
using System.Threading.Tasks;
using System.IO.Pipes;
using TimeoutException = System.TimeoutException;

namespace VM_SelfManager
{
    class VM_SelfManager : ServiceBase
    {
        
        /// <summary>
        /// Static class acting as a container of methods found useful in manipulating WMI data.
        /// </summary>
        /// <remarks>
        /// This class exists only to make it easier for me to be lazy.
        /// 
        /// Most items in this class are directly copied or adapted from the public documentation at:
        ///    http://msdn.microsoft.com/en-us/library/cc723869(v=VS.85).aspx
        /// </remarks>
        protected static class Utility
        {
            public static class ResourceType
            {
                public const UInt16 Other = 1;
                public const UInt16 ComputerSystem = 2;
                public const UInt16 Processor = 3;
                public const UInt16 Memory = 4;
                public const UInt16 IDEController = 5;
                public const UInt16 ParallelSCSIHBA = 6;
                public const UInt16 FCHBA = 7;
                public const UInt16 iSCSIHBA = 8;
                public const UInt16 IBHCA = 9;
                public const UInt16 EthernetAdapter = 10;
                public const UInt16 OtherNetworkAdapter = 11;
                public const UInt16 IOSlot = 12;
                public const UInt16 IODevice = 13;
                public const UInt16 FloppyDrive = 14;
                public const UInt16 CDDrive = 15;
                public const UInt16 DVDdrive = 16;
                public const UInt16 Serialport = 17;
                public const UInt16 Parallelport = 18;
                public const UInt16 USBController = 19;
                public const UInt16 GraphicsController = 20;
                public const UInt16 StorageExtent = 21;
                public const UInt16 Disk = 22;
                public const UInt16 Tape = 23;
                public const UInt16 OtherStorageDevice = 24;
                public const UInt16 FirewireController = 25;
                public const UInt16 PartitionableUnit = 26;
                public const UInt16 BasePartitionableUnit = 27;
                public const UInt16 PowerSupply = 28;
                public const UInt16 CoolingDevice = 29;
                public const UInt16 DisketteController = 1;
           }
            public static class ReturnCode
            {
                public const UInt32 Completed = 0;
                public const UInt32 Started = 4096;
                public const UInt32 Failed = 32768;
                public const UInt32 AccessDenied = 32769;
                public const UInt32 NotSupported = 32770;
                public const UInt32 Unknown = 32771;
                public const UInt32 Timeout = 32772;
                public const UInt32 InvalidParameter = 32773;
                public const UInt32 SystemInUse = 32774;
                public const UInt32 InvalidState = 32775;
                public const UInt32 IncorrectDataType = 32776;
                public const UInt32 SystemNotAvailable = 32777;
                public const UInt32 OutofMemory = 32778;
            }
            public static class JobState
            {
                public const UInt16 New = 2;
                public const UInt16 Starting = 3;
                public const UInt16 Running = 4;
                public const UInt16 Suspended = 5;
                public const UInt16 ShuttingDown = 6;
                public const UInt16 Completed = 7;
                public const UInt16 Terminated = 8;
                public const UInt16 Killed = 9;
                public const UInt16 Exception = 10;
                public const UInt16 Service = 11;
            }
            public static class VMState
            {
                public const int Paused = 32768,
                                 Stopped = 3,
                                 Running = 2,
                                 Saving = 32773,
                                 Stopping = 32774,
                                 Snapshotting = 32771,
                                 Suspended = 32769,
                                 Starting = 32770;
            }

            
            public static ManagementObject GetVM(string VMName)
            {
                ObjectQuery query = new ObjectQuery("select * from MSVM_ComputerSystem where ElementName like '"+VMName+"'");
                ManagementObjectSearcher searcher = new ManagementObjectSearcher(new ManagementScope(@"root\virtualization",null), query);
                ManagementObjectCollection VMs = searcher.Get();
                if (VMs.Count != 1)
                    return null;
                return VMs.Cast<ManagementObject>().FirstOrDefault();
            }

            public static bool JobCompleted(ManagementBaseObject outParams)
            {
                bool jobCompleted = true;

                //Retrieve msvc_StorageJob path. This is a full wmi path
                string JobPath = (string)outParams["Job"];
                ManagementObject Job = new ManagementObject(new ManagementScope(@"root\virtualization"), new ManagementPath(JobPath), null);
                //Try to get storage job information
                Job.Get();
                UInt16 jobState = (UInt16)Job["JobState"];

                return jobState == JobState.Completed;
            }

            public static bool JobFailed(ManagementBaseObject outParams)
            {
                bool jobCompleted = true;

                //Retrieve msvc_StorageJob path. This is a full wmi path
                string JobPath = (string)outParams["Job"];
                ManagementObject Job = new ManagementObject(new ManagementScope(@"root\virtualization"), new ManagementPath(JobPath), null);
                //Try to get storage job information
                Job.Get();
                UInt16 jobState = (UInt16)Job["JobState"];

                return (jobState > JobState.Completed && jobState < JobState.Service);
            }

            public static bool WaitForJob(ManagementBaseObject outParams)
            {
                bool JobSuccess = true;
                string JobPath = (string)outParams["Job"];
                ManagementObject Job = new ManagementObject(new ManagementScope(@"root\virtualization"), new ManagementPath(JobPath), null);
                //Try to get storage job information
                Job.Get();
                while ((UInt16)Job["JobState"] == JobState.Starting
                    || (UInt16)Job["JobState"] == JobState.Running)
                {
                    Console.WriteLine("In progress... {0}% completed.", Job["PercentComplete"]);
                    System.Threading.Thread.Sleep(1000);
                    Job.Get();
                }
                //Figure out if job failed
                UInt16 jobState = (UInt16)Job["JobState"];
                if (jobState != JobState.Completed)
                {
                    JobSuccess = false;
                }
                return JobSuccess;
            }

            public static ManagementObjectCollection GetActiveVMs()
            {
                ObjectQuery query = new ObjectQuery("select * from MSVM_ComputerSystem where ProcessID > 0");
                ManagementObjectSearcher searcher = new ManagementObjectSearcher(new ManagementScope(@"root\virtualization", null), query);
                return searcher.Get();
            }



        }

        protected static class Process
        {
            public static void NewSnapshot(string VMName, string SnapshotName = null)
            {
                ManagementClass VSMS = new ManagementClass(new ManagementScope(@"root\virtualization",null).Path.Path, "MSVM_VirtualSystemManagementService", null);

            }
        }

        protected class Listener : IDisposable
        {

            public Listener(string name, string pipe, CancellationTokenSource tokenSource, WriterDelegate writer = null)
            {
                VMName = name;
                Pipe = pipe;
                TokenSource = tokenSource;
                Writer = writer;
                Task = new Task(() =>
                                    {
                                        NamedPipeClientStream pipeStream = new NamedPipeClientStream(Pipe);
                                        while (!Task.IsCanceled)
                                            Listen(pipeStream, TokenSource.Token);
                                    }, TokenSource.Token);
            }

            public string VMName;
            public string Pipe;
            public Task Task;
            public CancellationTokenSource TokenSource;
            public delegate void WriterDelegate(string Message, EventLogEntryType EventType, int ID, short Category);
            public WriterDelegate Writer;

            public void Dispose()
            {
                TokenSource.Cancel();
            }

            protected void write(string message, EventLogEntryType type = EventLogEntryType.Information, int ID = 0, short category = (short)0)
            {
                if (Writer != null)
                    Writer(message, type, ID, category);
            }

            protected static class ProcessingKeys
            {
                public const byte NewSnapshot = 0x0E; // Crtl-N
                public const byte OpenSnapshot = 0x0F; // Ctrl-O
                public const byte WriteToLog = 0x17; // Ctrl-W
                public const byte ReadFromLog = 0x12; // Ctrl-R
                public const byte BootVM = 0x02; // Ctrl-B
                public const byte EndMessage = 0x05; // Crtl-E
            }

            protected void Listen(NamedPipeClientStream pipe, CancellationToken token)
            {
                byte current = (byte) pipe.ReadByte();
                while (current != -1 && !token.IsCancellationRequested)
                {
                    //ProcessMessage(current, pipe, token);
                    if (token.IsCancellationRequested)
                        return;

                    switch (current)
                    {
                        case ProcessingKeys.BootVM:
                            ProcessBootVM(pipe, token);
                            break;
                        case ProcessingKeys.NewSnapshot:
                            ProcessNewSnapshot(pipe, token);
                            break;
                        case ProcessingKeys.OpenSnapshot:
                            ProcessOpenSnapshot(pipe, token);
                            break;
                        case ProcessingKeys.WriteToLog:
                            ProcessWriteToLog(pipe, token);
                            break;
                        case ProcessingKeys.ReadFromLog:
                            ProcessReadFromLog(pipe, token);
                            break;
                    }

                    current = (byte) pipe.ReadByte();
                }

            }

            private string ReadToEnd(NamedPipeClientStream pipe)
            {
                List<byte> input = new List<byte>();
                byte c = (byte)pipe.ReadByte();
                while (c != ProcessingKeys.EndMessage)
                {
                    input.Add(c);
                    c = (byte)pipe.ReadByte();
                }
                return new string(input.Cast<char>().ToArray());
            }

            protected void ProcessBootVM(NamedPipeClientStream pipe, CancellationToken token)
            {
                //Get the name of the VM in question
                string name = ReadToEnd(pipe);
                ManagementObject VM = Utility.GetVM(name);
                ManagementBaseObject input = VM.GetMethodParameters("RequestStateChange");
                input["RequestedState"] = Utility.VMState.Running;
                write("Starting VM '"+name+"'.  Requested by '"+VMName);
                if (Utility.WaitForJob(VM.InvokeMethod("RequestStateChange", input, null)))
                    write("Started VM '" + name + "' successfully.  Requested by '" + VMName);
                else
                    write("Failed to start VM '" + name + "', as requested by '" + VMName, EventLogEntryType.Error);
            }

            protected void ProcessNewSnapshot(NamedPipeClientStream pipe, CancellationToken token)
            {
                //Get name for snapshot (if any)
                string name = ReadToEnd(pipe);
                ManagementClass VSMS = new ManagementClass(new ManagementScope(@"root\virtualization", null).Path.Path, "MSVM_VirtualSystemManagementService", null);
                ManagementBaseObject SnapshotInput = VSMS.GetMethodParameters("CreateVirtualSystemSnapshot");
                SnapshotInput["SourceSystem"] = Utility.GetVM(VMName);
                ManagementObject returned;
                SnapshotInput["SnapshotSettingData"]
                VSMS.InvokeMethod();

                if (name != null)
                {
                    ManagementBaseObject RenameInput = VSMS.GetMethodParameters("ModifyVirtualSystem");

                }
            }

            protected void ProcessOpenSnapshot(NamedPipeClientStream pipe, CancellationToken token)
            {
            }
            
            protected void ProcessWriteToLog(NamedPipeClientStream pipe, CancellationToken token)
            {
            }
            
            protected void ProcessReadFromLog(NamedPipeClientStream pipe, CancellationToken token)
            {
            }
        }

        protected Dictionary<string, Listener> Connections;
        protected Timer UpdateTimer;

        public VM_SelfManager()
        {
            ServiceName = "VM_SelfManager";
            EventLog.Log = "Application";
            EventLog.Source = ServiceName;
            // Events to enable
            CanHandlePowerEvent = false;
            CanHandleSessionChangeEvent = false;
            CanPauseAndContinue = true;
            CanShutdown = true;
            CanStop = true;

            Connections = new Dictionary<string, Listener>();
            UpdateTimer = new Timer(UpdateWorld, null, 0, (60*1000)); // 60 second timer
        }

        static void Main()
        {
            ServiceBase.Run(new VM_SelfManager());
        }

        #region Standard Service Control Methods

        protected override void Dispose(bool disposing)
        {
            base.Dispose(disposing);
        }

        /// <summary>
        /// OnStart(): Put startup code here
        ///  - Start threads, get inital data, etc.
        /// </summary>
        /// <param name="args"></param>
        protected override void OnStart(string[] args)
        {
            base.OnStart(args);
        }

        /// <summary>
        /// OnStop(): Put your stop code here
        /// - Stop threads, set final data, etc.
        /// </summary>
        protected override void OnStop()
        {
            base.OnStop();
        }

        /// <summary>
        /// OnPause: Put your pause code here
        /// - Pause working threads, etc.
        /// </summary>
        protected override void OnPause()
        {
            base.OnPause();
        }

        /// <summary>
        /// OnContinue(): Put your continue code here
        /// - Un-pause working threads, etc.
        /// </summary>
        protected override void OnContinue()
        {
            base.OnContinue();
        }

        /// <summary>
        /// OnShutdown(): Called when the System is shutting down
        /// - Put code here when you need special handling
        ///   of code that deals with a system shutdown, such
        ///   as saving special data before shutdown.
        /// </summary>
        protected override void OnShutdown()
        {
            base.OnShutdown();
        }

        #endregion

        protected void WriteEvent(string Message, EventLogEntryType EventType, int ID, short Category)
        {
            EventLog.WriteEntry(Message, EventType, ID, Category);
        }

        protected void UpdateWorld(object NotUsed)
        {
            Dictionary<string, Listener> tempConn = new Dictionary<string, Listener>();
            foreach (ManagementObject VM in Utility.GetActiveVMs()) //grab all active VMs
            {
                foreach (ManagementObject SP in VM.GetRelated("MSVM_SerialPort")) // grab all serial ports on each VM
                    if (SP["ElementName"].ToString().Equals("COM 2", StringComparison.CurrentCultureIgnoreCase)) // only care about COM2 for each VM
                    {
                        string con = SP.GetRelated("MSVM_ResourceAllocationSettingData").Cast<ManagementObject>().FirstOrDefault()["Coonection"].ToString();
                        if (!(con.Equals(String.Empty))) // Is there a pipe connection listed for this port?
                        {
                            if (tempConn.ContainsKey(con))
                            {
                                /* I don't expect this condition to happen unless someone was dumb
                                 * This covers the case where someone has assigned the same pipe to two seperate VMs.
                                 * If this occurs, only the first to be recorded will be processed.  All others will
                                 *   be discarded.
                                 */
                            }
                            if (Connections.ContainsKey(con))
                            {
                                // If already tracking, transfer info to new connection list and remove from old list.
                                tempConn.Add(con, Connections[con]);
                                Connections.Remove(con);
                            }
                            else
                            {
                                // Not already being tracked: start a listener and track it.
                                tempConn.Add(con, new Listener(VM["ElementName"].ToString(), con, new CancellationTokenSource()));
                            }

                        }
                    }
            }

            foreach (var connection in Connections)
            {
                //kill any connections that no longer have active VMs to talk to
                connection.Value.Dispose();
            }
            //replace the current active dictionary with the one we've been filling
            Connections = tempConn;
        }

    }
}
