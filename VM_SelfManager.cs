﻿/**
  *    Copyright 2012 Tim Rogers
  *
  *   Licensed under the Apache License, Version 2.0 (the "License");
  *   you may not use this file except in compliance with the License.
  *   You may obtain a copy of the License at
  *
  *       http://www.apache.org/licenses/LICENSE-2.0
  *
  *   Unless required by applicable law or agreed to in writing, software
  *   distributed under the License is distributed on an "AS IS" BASIS,
  *   WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
  *   See the License for the specific language governing permissions and
  *   limitations under the License.
  *
  */

using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.ServiceProcess;
using System.Management;
using System.Threading;
using System.Threading.Tasks;
using System.IO.Pipes;
using System.IO;
using Microsoft.Win32;

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

            /// <summary>
            /// Common utility function to get a service object
            /// </summary>
            /// <param name="scope"></param>
            /// <param name="serviceName"></param>
            /// <returns></returns>
            public static ManagementObject GetServiceObject(string serviceName)
            {
                ManagementScope scope = new ManagementScope(@"root\virtualization", null);
                scope.Connect();
                ManagementPath wmiPath = new ManagementPath(serviceName);
                ManagementClass serviceClass = new ManagementClass(scope, wmiPath, null);
                ManagementObjectCollection services = serviceClass.GetInstances();

                ManagementObject serviceObject = null;

                foreach (ManagementObject service in services)
                {
                    serviceObject = service;
                }
                return serviceObject;
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

            public static ManagementObject GetSnapshot(string VMName, string SnapshotName = null)
            {
                ObjectQuery query = new ObjectQuery("select * from MSVM_ComputerSystem where ElementName like '" + VMName + "'");
                ManagementObjectSearcher searcher = new ManagementObjectSearcher(new ManagementScope(@"root\virtualization", null), query);
                ManagementObjectCollection VMs = searcher.Get();
                if (VMs.Count != 1)
                    return null;
                ManagementObject VM = VMs.Cast<ManagementObject>().FirstOrDefault();
                ManagementObjectCollection Snaps = VM.GetRelated("MSVM_VirtualSystemSettingData");
                
                if (SnapshotName != null) // Looking for a specific snapshot
                {
                    foreach (ManagementObject snap in Snaps)
                    {
                        if (snap["ElementName"].ToString().Equals(SnapshotName,
                                                                  StringComparison.CurrentCultureIgnoreCase))
                            return snap;
                    }
                    return null;
                }
                // Looking for the most recent snapshot
                string parent = null;
                foreach (ManagementObject snap in Snaps)
                {
                    if (snap["ElementName"].ToString().Equals(VMName))
                        parent = snap["Parent"].ToString();
                }
                
                if (parent == null || parent.Equals(String.Empty))
                {
                    return null;
                }
                int startpoint = parent.IndexOf(".InstanceID=") + 13;
                string parentInstance = parent.Substring(startpoint, parent.Length - (startpoint + 1));

                foreach (ManagementObject snap in Snaps)
                {
                    if (snap["InstanceID"].ToString().Equals(parentInstance, StringComparison.CurrentCultureIgnoreCase))
                        return snap;
                }
                return null;
            }

            public static bool JobCompleted(ManagementBaseObject outParams)
            {
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
                try
                {
                    bool JobSuccess = true;
                    string JobPath = (string)outParams["Job"];
                    if (JobPath == null)
                    {
                        // No path to the job object.  Assume that either it is done already or no job was produced.
                        return JobSuccess;
                    }
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
                catch (Exception E)
                {
                    return false;
                }
            }

            public static ManagementObjectCollection GetActiveVMs()
            {
                ObjectQuery query = new ObjectQuery("select * from MSVM_ComputerSystem where ProcessID > 0");
                ManagementObjectSearcher searcher = new ManagementObjectSearcher(new ManagementScope(@"root\virtualization", null), query);
                return searcher.Get();
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
                                        string trimmed = Pipe.Substring(2); // trim the leading '\\'
                                        const string PipeID = @"\pipe\";
                                        int index = trimmed.IndexOf(PipeID);
                                        string pipeServer = trimmed.Substring(0, index);
                                        string pipeName = trimmed.Substring(index + PipeID.Length);
                                        NamedPipeClientStream pipeStream = new NamedPipeClientStream(pipeServer, pipeName);
                                        try { 
                                            pipeStream.Connect(3000); 
                                        }
                                        catch (Exception)
                                        {
                                            write("Failed to make connection to pipe: " + Pipe, EventLogEntryType.Error);
                                            this.Dispose();
                                            return;
                                        }
                                        while (!Task.IsCanceled && !TokenSource.IsCancellationRequested)
                                            Listen(pipeStream, TokenSource.Token);
                                        try{pipeStream.Close();}
                                        catch (Exception){} // I don't really care if closing the pipe failed, I just don't want to kill the program.
                                    }, TokenSource.Token);
                Task.Start();
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
                public const byte DeleteSnapshot = 0x04; // Ctrl-D
            }

            protected void Listen(NamedPipeClientStream pipe, CancellationToken token)
            {
                byte current = (byte) pipe.ReadByte();
                while (current != 255 && !token.IsCancellationRequested)
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
                        case ProcessingKeys.DeleteSnapshot:
                            ProcessDeleteSnapshot(pipe, token);
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
                if (current == 255)
                    this.Dispose();
            }

            private static string ReadToEnd(NamedPipeClientStream pipe)
            {
                List<char> input = new List<char>();
                char c = (char)pipe.ReadByte();
                while (c != ProcessingKeys.EndMessage)
                {
                    input.Add(c);
                    c = (char)pipe.ReadByte();
                }
                if (input.Count > 0)
                    return new string(input.ToArray());
                return null;
            }

            private static byte[] ReadBytesToEnd(NamedPipeClientStream pipe)
            {
                List<byte> input = new List<byte>();
                byte c = (byte)pipe.ReadByte();
                while (c != ProcessingKeys.EndMessage)
                {
                    input.Add(c);
                    c = (byte)pipe.ReadByte();
                }
                if (input.Count > 0)
                    return input.ToArray();
                return null;
            }

            protected void ProcessBootVM(NamedPipeClientStream pipe, CancellationToken token)
            {
                //Get the name of the VM in question
                string name = ReadToEnd(pipe);
                //short circut if no name provided
                if (null == name)
                    return;
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
                try
                {
                    //Get name for snapshot (if any)
                    string name = ReadToEnd(pipe);
                    ManagementObject VSMS = Utility.GetServiceObject("MSVM_VirtualSystemManagementService");
                        //new ManagementClass(new ManagementScope(@"root\virtualization", null).Path.Path, "MSVM_VirtualSystemManagementService", null);
                    ManagementBaseObject SnapshotInput = VSMS.GetMethodParameters("CreateVirtualSystemSnapshot");
                    ManagementObject VM = Utility.GetVM(VMName);
                    SnapshotInput["SourceSystem"] = VM.Path.Path;
                    //ManagementObject returned = new ManagementObject();
                    //SnapshotInput["SnapshotSettingData"] = returned;
                    ManagementBaseObject WorkJob = VSMS.InvokeMethod("CreateVirtualSystemSnapshot", SnapshotInput, null);
                    bool success = Utility.WaitForJob(WorkJob);
                    if (!success)
                        throw new Exception();
                    // This is a kludge because I cannot get 'CreateVirtualSystenSnapshot' to return a non-null 
                    //   SnapshotSettingData object for me, so I'm just grabbing the snapshot I just made for myself.
                    ManagementObject returned = Utility.GetSnapshot(VMName);
                    string OldName = returned["ElementName"].ToString();
                    write("Snapshot created successfully:  " + VMName + "::" + OldName);

                    if (name != null)
                    {
                        try
                        {
                            // There is a hard limit of 100 characters for a snapshot name.  This limit is imposed by Hyper-V.
                            if (name.Length > 100)
                                name = name.Substring(0, 100);
                            returned["ElementName"] = name;
                            ManagementBaseObject RenameInput = VSMS.GetMethodParameters("ModifyVirtualSystem");
                            RenameInput["ComputerSystem"] = VM.Path.Path;
                            RenameInput["SystemSettingData"] = returned.GetText(TextFormat.CimDtd20);
                            WorkJob = VSMS.InvokeMethod("ModifyVirtualSystem", RenameInput, null);
                            success = Utility.WaitForJob(WorkJob);
                            if (!success)
                                throw new Exception();
                            write("Snapshot renamed successfully:  " + OldName + " ==> " + name);
                        }
                        catch (Exception)
                        {
                            write("Error renaming snapshot:  " + OldName + " ==> " + name, EventLogEntryType.Error);
                        }
                    }
                }
                catch (Exception E)
                {
                    write("Snapshot creation failed on VM: " + VMName);
                }
            }

            protected void ProcessOpenSnapshot(NamedPipeClientStream pipe, CancellationToken token)
            {
                //Get name for snapshot (if any)
                string name = ReadToEnd(pipe);
                ManagementObject VSMS = Utility.GetServiceObject("MSVM_VirtualSystemManagementService");
                //ManagementClass VSMS = new ManagementClass(new ManagementScope(@"root\virtualization", null).Path.Path, "MSVM_VirtualSystemManagementService", null);
                ManagementObject Snapshot = Utility.GetSnapshot(VMName, name);
                if (Snapshot == null)
                    return;
                // Using 'ApplyVirtualSystemSnapshotEx' instead of 'ApplyVirtualSystemSnapshot' because this one returns a Job object.
                ManagementBaseObject input = VSMS.GetMethodParameters("ApplyVirtualSystemSnapshotEx");
                ManagementObject VM = Utility.GetVM(VMName);
                input["ComputerSystem"] = VM.Path.Path;
                input["SnapshotSettingData"] = Snapshot.Path.Path;
                // Having now set up all parameters, we need to stop the VM before we can apply the snapshot.
                ManagementBaseObject stopInput = VM.GetMethodParameters("RequestStateChange");
                ManagementBaseObject startInput = VM.GetMethodParameters("RequestStateChange");
                stopInput["RequestedState"] = Utility.VMState.Stopped;
                startInput["RequestedState"] = Utility.VMState.Running;
                if (!Utility.WaitForJob(VM.InvokeMethod("RequestStateChange", stopInput, null)))
                {
                    write("Apply Snapshot failed.  Reason: Unable to stop VM '" + VMName + "'.", EventLogEntryType.Error);
                    return;
                }
                ManagementBaseObject output = VSMS.InvokeMethod("ApplyVirtualSystemSnapshotEx", input, null);
                bool success = Utility.WaitForJob(output);
                if (success)
                    write("Applied snapshot to VM:  " + (name ?? "current") + " ==> " + VMName);
                else
                {
                    write("Failed to apply snapshot to VM:  " + (name ?? "current") + " ==> " + VMName, EventLogEntryType.Error);
                    return;
                }
                if (!Utility.WaitForJob(VM.InvokeMethod("RequestStateChange", startInput, null)))
                {
                    write("Unable to start VM '" + VMName + "' after applying snapshot.", EventLogEntryType.Error);
                }
                
            }

            protected void ProcessDeleteSnapshot(NamedPipeClientStream pipe, CancellationToken token)
            {
                // Are we deleting a whole tree?
                char c = (char)pipe.ReadByte();
                if (c == ProcessingKeys.DeleteSnapshot)
                {
                    ProcessDeleteTree(pipe, token);
                    return;
                }
                //Get name for snapshot (if any)
                string name = (c != ProcessingKeys.EndMessage ? String.Concat(c, ReadToEnd(pipe)) : null);
                ManagementObject VSMS = Utility.GetServiceObject("MSVM_VirtualSystemManagementService");
                //ManagementClass VSMS = new ManagementClass(new ManagementScope(@"root\virtualization", null).Path.Path, "MSVM_VirtualSystemManagementService", null);
                ManagementObject Snapshot = Utility.GetSnapshot(VMName, name);
                if (Snapshot == null)
                    return;
                ManagementBaseObject input = VSMS.GetMethodParameters("RemoveVirtualSystemSnapshot");
                input["SnapshotSettingData"] = Snapshot.Path.Path;

                if (Utility.WaitForJob(VSMS.InvokeMethod("RemoveVirtualSystemSnapshot", input, null)))
                    write("Deleted snapshot:  " + VMName + "::" + (name ?? "current"));
                else
                    write("Failed to delete snapshot:  " + VMName + "::" + (name ?? "current"), EventLogEntryType.Error);
            }

            protected void ProcessDeleteTree(NamedPipeClientStream pipe, CancellationToken token)
            {
                //Get name for snapshot (if any)
                string name = ReadToEnd(pipe);
                ManagementObject VSMS = Utility.GetServiceObject("MSVM_VirtualSystemManagementService");
                //ManagementClass VSMS = new ManagementClass(new ManagementScope(@"root\virtualization", null).Path.Path, "MSVM_VirtualSystemManagementService", null);
                ManagementObject Snapshot = Utility.GetSnapshot(VMName, name);
                if (Snapshot == null)
                    return;
                ManagementBaseObject input = VSMS.GetMethodParameters("RemoveVirtualSystemSnapshotTree");
                input["SnapshotSettingData"] = Snapshot.Path.Path;

                if (Utility.WaitForJob(VSMS.InvokeMethod("RemoveVirtualSystemSnapshotTree", input, null)))
                    write("Deleted snapshot tree:  " + VMName + "::" + (name ?? "current"));
                else
                    write("Failed to delete snapshot tree:  " + VMName + "::" + (name ?? "current"), EventLogEntryType.Error);
            }

            protected void ProcessWriteToLog(NamedPipeClientStream pipe, CancellationToken token)
            {
                string path = Path.Combine(VM_SelfManager.LogPath, VMName + ".log");
                FileStream Log = new FileStream(path, FileMode.Append,FileAccess.Write);
                byte[] Data = ReadBytesToEnd(pipe);
                Log.Write(Data, 0, Data.Length);
                Log.Close();
            }
            
            protected void ProcessReadFromLog(NamedPipeClientStream pipe, CancellationToken token)
            {
                //Just doing this to clear the stream for the next command.
                ReadToEnd(pipe);
                string path = Path.Combine(VM_SelfManager.LogPath, VMName + ".log");
                FileStream Log = new FileStream(path, FileMode.OpenOrCreate, FileAccess.Read);
                
                // This is only complicated in order to handle cases of WAY oversized log files.
                int reps = (int)(Log.Length / int.MaxValue);
                byte[] data = new byte[(reps==0 ? (int)(Log.Length % int.MaxValue) : int.MaxValue)];
                
                //read the log
                Log.Seek(0, SeekOrigin.Begin);
                for (int i = 0; i <= reps; i++)
                {
                    Log.Read(data, 0, (i == reps ? (int) (Log.Length%int.MaxValue) : int.MaxValue));
                    //write it immediately so we can grab the next set if needed...
                    pipe.Write(data, 0, (i == reps ? (int) (Log.Length%int.MaxValue) : int.MaxValue));
                }
                pipe.WriteByte(0x1A); // writes an 'end of stream' character so the client knows that all has been sent.
            }
        }

        protected Dictionary<string, Listener> Connections;
        protected Timer UpdateTimer;
        protected static string LogPath;
        protected bool LockObject;
        protected const int TimerReset = 10 * 1000;

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
        }

        static void Main()
        {
            /*
            VM_SelfManager Manager = new VM_SelfManager();
            Manager.OnStart(new string[0]);
            ConsoleKeyInfo c = new ConsoleKeyInfo(' ',ConsoleKey.Spacebar,false,false,false);
            while (!(c.Key.Equals(ConsoleKey.Escape)))
            {
                Console.Out.Write('.');
                System.Threading.Thread.Sleep(1000);
                if (Console.KeyAvailable)
                    c = Console.ReadKey(true);
            }
            Manager.OnStop();
             */
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
            LockObject = false;
            Connections = new Dictionary<string, Listener>();
            UpdateTimer = new Timer(UpdateWorld, null, 0, TimerReset); // 15 second timer
            LogPath = GetLogPath();

            base.OnStart(args);
        }

        /// <summary>
        /// OnStop(): Put your stop code here
        /// - Stop threads, set final data, etc.
        /// </summary>
        protected override void OnStop()
        {
            UpdateTimer.Dispose();
            foreach (KeyValuePair<string, Listener> pair in Connections)
            {
                pair.Value.Dispose();
            }
            Connections.Clear();

            base.OnStop();
        }

        /// <summary>
        /// OnPause: Put your pause code here
        /// - Pause working threads, etc.
        /// </summary>
        protected override void OnPause()
        {
            UpdateTimer.Change(TimeSpan.Zero,new TimeSpan(-1));
            base.OnPause();
        }

        /// <summary>
        /// OnContinue(): Put your continue code here
        /// - Un-pause working threads, etc.
        /// </summary>
        protected override void OnContinue()
        {
            UpdateTimer.Change(0, TimerReset);
            base.OnContinue();
        }

        #endregion

        protected void WriteEvent(string Message, EventLogEntryType EventType, int ID, short Category)
        {
            EventLog.WriteEntry(Message, EventType, ID, Category);
        }

        protected static string GetLogPath()
        {
            RegistryKey Base = Registry.LocalMachine.CreateSubKey(@"SOFTWARE\Hyper-V\VM Self Manager");
            if (Base == null)
                throw new Exception("Unable to open registry key.");
            string tempPath = (string)Base.GetValue("LogPath");
            if (tempPath != null)
                return tempPath;
            //else
            const string defautPath = @"C:\Hyper-V\Logs";
            Base.SetValue("LogPath",defautPath,RegistryValueKind.String);
            return defautPath;
        }

        protected void UpdateWorld(object NotUsed)
        {
            try
            {
                if (LockObject)
                    return;
                LockObject = true;
                foreach (var item in Connections)
                {
                    if (item.Value.TokenSource.IsCancellationRequested)
                        Connections.Remove(item.Key);
                }

                Dictionary<string, Listener> tempConn = new Dictionary<string, Listener>();
                foreach (ManagementObject VM in Utility.GetActiveVMs()) //grab all active VMs
                {
                    foreach (ManagementObject SP in VM.GetRelated("MSVM_SerialPort")) // grab all serial ports on each VM
                        if (SP["ElementName"].ToString().Equals("COM 2", StringComparison.CurrentCultureIgnoreCase)) // only care about COM2 for each VM
                        {
                            ManagementObjectCollection collection = SP.GetRelated("MSVM_ResourceAllocationSettingData");
                            if (collection.Count > 0)
                            {
                                ManagementObject item = null;
                                foreach (ManagementObject o in collection)
                                {
                                    item = o;
                                    break;
                                }
                                string con;
                                if (item != null)
                                    con = ((string[])item["Connection"])[0];
                                else con = String.Empty;

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
                                        tempConn.Add(con, new Listener(VM["ElementName"].ToString(), con, new CancellationTokenSource(), WriteEvent));
                                    }

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
            catch (InvalidOperationException E)
            {
                // I'm going to ignore this because it most likely means that a VM turned off during enumeration.
                // But I do want the timer to re-run this updater to catch the changes.
                UpdateTimer.Change(500, TimerReset);
            }
            catch (Exception E)
            {
                WriteEvent(E.ToString(), EventLogEntryType.Error, 0, 0);
            }
            finally
            {
                LockObject = false;
            }
        }
    }
}
