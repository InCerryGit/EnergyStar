using EnergyStar.Interop;
using System.Diagnostics;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Text;

namespace EnergyStar
{
    public unsafe class EnergyManager
    {
        public static readonly HashSet<string> BypassProcessList = new HashSet<string>
        {
#if DEBUG
            // Visual Studio
            "devenv.exe",
#endif
        };

        // Speical handling needs for UWP to get the child window process
        public const string UWPFrameHostApp = "ApplicationFrameHost.exe";

        private static uint pendingProcPid = 0;
        private static string pendingProcName = "";

        private static IntPtr pThrottleOn = IntPtr.Zero;
        private static IntPtr pThrottleOff = IntPtr.Zero;
        private static int szControlBlock = 0;

        static EnergyManager()
        {
            var settings = Settings.Load();
            BypassProcessList.UnionWith(settings.Exemptions.Select(x => x.ToLowerInvariant()));

            szControlBlock = Marshal.SizeOf<Win32Api.PROCESS_POWER_THROTTLING_STATE>();
            pThrottleOn = Marshal.AllocHGlobal(szControlBlock);
            pThrottleOff = Marshal.AllocHGlobal(szControlBlock);

            var throttleState = new Win32Api.PROCESS_POWER_THROTTLING_STATE
            {
                Version = Win32Api.PROCESS_POWER_THROTTLING_STATE.PROCESS_POWER_THROTTLING_CURRENT_VERSION,
                ControlMask = Win32Api.ProcessorPowerThrottlingFlags.PROCESS_POWER_THROTTLING_EXECUTION_SPEED,
                StateMask = Win32Api.ProcessorPowerThrottlingFlags.PROCESS_POWER_THROTTLING_EXECUTION_SPEED,
            };

            var unthrottleState = new Win32Api.PROCESS_POWER_THROTTLING_STATE
            {
                Version = Win32Api.PROCESS_POWER_THROTTLING_STATE.PROCESS_POWER_THROTTLING_CURRENT_VERSION,
                ControlMask = Win32Api.ProcessorPowerThrottlingFlags.PROCESS_POWER_THROTTLING_EXECUTION_SPEED,
                StateMask = Win32Api.ProcessorPowerThrottlingFlags.None,
            };

            Marshal.StructureToPtr(throttleState, pThrottleOn, false);
            Marshal.StructureToPtr(unthrottleState, pThrottleOff, false);
        }

        private static void ToggleEfficiencyMode(IntPtr hProcess, bool enable)
        {
            Win32Api.SetProcessInformation(hProcess, Win32Api.PROCESS_INFORMATION_CLASS.ProcessPowerThrottling,
                enable ? pThrottleOn : pThrottleOff, (uint) szControlBlock);
            Win32Api.SetPriorityClass(hProcess,
                enable ? Win32Api.PriorityClass.IDLE_PRIORITY_CLASS : Win32Api.PriorityClass.NORMAL_PRIORITY_CLASS);
        }

        private static string GetProcessNameFromHandle(IntPtr hProcess)
        {
            int capacity = 1024;
            var sb = new StringBuilder(capacity);

            if (Win32Api.QueryFullProcessImageName(hProcess, 0, sb, ref capacity))
            {
                return Path.GetFileName(sb.ToString());
            }

            return "";
        }

        public static unsafe void HandleForegroundEvent(IntPtr hwnd)
        {
            var windowThreadId = Win32Api.GetWindowThreadProcessId(hwnd, out uint procId);
            // This is invalid, likely a process is dead, or idk
            if (windowThreadId == 0 || procId == 0) return;

            var procHandle = Win32Api.OpenProcess(
                (uint) (Win32Api.ProcessAccessFlags.QueryLimitedInformation |
                        Win32Api.ProcessAccessFlags.SetInformation), false, procId);
            if (procHandle == IntPtr.Zero) return;

            // Get the process
            var appName = GetProcessNameFromHandle(procHandle);

            // UWP needs to be handled in a special case
            if (appName == UWPFrameHostApp)
            {
                var found = false;
                Win32Api.EnumChildWindows(hwnd, (innerHwnd, lparam) =>
                {
                    if (found) return true;
                    if (Win32Api.GetWindowThreadProcessId(innerHwnd, out uint innerProcId) > 0)
                    {
                        if (procId == innerProcId) return true;

                        var innerProcHandle = Win32Api.OpenProcess(
                            (uint) (Win32Api.ProcessAccessFlags.QueryLimitedInformation |
                                    Win32Api.ProcessAccessFlags.SetInformation), false, innerProcId);
                        if (innerProcHandle == IntPtr.Zero) return true;

                        // Found. Set flag, reinitialize handles and call it a day
                        found = true;
                        Win32Api.CloseHandle(procHandle);
                        procHandle = innerProcHandle;
                        procId = innerProcId;
                        appName = GetProcessNameFromHandle(procHandle);
                    }

                    return true;
                }, IntPtr.Zero);
            }

            // Boost the current foreground app, and then impose EcoQoS for previous foreground app
            var bypass = BypassProcessList.Contains(appName.ToLowerInvariant());
            if (!bypass)
            {
                var currentHandle = procHandle;
                var currentName = appName;
                var currentSessionId = Process.GetCurrentProcess().SessionId;
                var runningProcesses = Process.GetProcesses().Where(p => p.SessionId == currentSessionId).ToDictionary(p => p.Id, p => p.ProcessName);
                do
                {
                    ToggleEfficiencyMode(currentHandle, false);
                    Console.WriteLine($"Boost {currentName}");

                    if (currentHandle != procHandle) Win32Api.CloseHandle(currentHandle);
                    
                    var status = Win32Api.NtQueryInformationProcess(currentHandle, 0, out var processInformation,
                        Unsafe.SizeOf<Win32Api.ParentProcessUtilities>(), out _);
                    if (status != 0 ||
                        processInformation.InheritedFromUniqueProcessId == IntPtr.Zero ||
                        runningProcesses.ContainsKey((int) processInformation.InheritedFromUniqueProcessId) == false)
                        break;
                    
                    currentHandle = Win32Api.OpenProcess(
                        (uint) (Win32Api.ProcessAccessFlags.QueryLimitedInformation |
                                Win32Api.ProcessAccessFlags.SetInformation), false, (uint)processInformation.InheritedFromUniqueProcessId);

                    currentName = GetProcessNameFromHandle(currentHandle);
                    if (BypassProcessList.Contains(currentName.ToLowerInvariant()))
                    {
                        break;
                    }
                    
                } while (true);
            }

            if (pendingProcPid != 0)
            {
                Console.WriteLine($"Throttle {pendingProcName}");

                var prevProcHandle = Win32Api.OpenProcess((uint) Win32Api.ProcessAccessFlags.SetInformation, false,
                    pendingProcPid);
                if (prevProcHandle != IntPtr.Zero)
                {
                    ToggleEfficiencyMode(prevProcHandle, true);
                    Win32Api.CloseHandle(prevProcHandle);
                    pendingProcPid = 0;
                    pendingProcName = "";
                }
            }

            if (!bypass)
            {
                pendingProcPid = procId;
                pendingProcName = appName;
            }

            Win32Api.CloseHandle(procHandle);
        }

        public static void ThrottleAllUserBackgroundProcesses()
        {
            var runningProcesses = Process.GetProcesses();
            var currentSessionId = Process.GetCurrentProcess().SessionId;

            var sameAsThisSession = runningProcesses.Where(p => p.SessionId == currentSessionId).ToDictionary(p => (uint)p.Id, p => p);
            foreach (var (_, proc) in sameAsThisSession)
            {
                if (proc.Id == pendingProcPid) continue;
                var processNameLower = $"{proc.ProcessName}.exe".ToLowerInvariant();
                if (BypassProcessList.Contains(processNameLower)) continue;
                var hProcess = Win32Api.OpenProcess(
                    (uint) (Win32Api.ProcessAccessFlags.SetInformation | Win32Api.ProcessAccessFlags.QueryInformation),
                    false, (uint) proc.Id);

                var needToggle = true;
                var parentProcesses = GetParentProcessNames(hProcess);
                foreach (var (parentProcId, parentProcName) in parentProcesses)
                {
                    if (parentProcId == pendingProcPid || BypassProcessList.Contains(parentProcName))
                    {
                        needToggle = false;
                        break;
                    }

                    if (sameAsThisSession.ContainsKey(parentProcId) == false)
                    {
                        break;
                    }
                }

                if (needToggle)
                {
                    ToggleEfficiencyMode(hProcess, true);    
                }
                
                Win32Api.CloseHandle(hProcess);
            }
        }

        public static IEnumerable<(uint ProcessId, string ProcessName)> GetParentProcessNames(IntPtr hProcess)
        {
            var current = hProcess;
            do
            {
                var status = Win32Api.NtQueryInformationProcess(current, 0, out var processInformation,
                    Unsafe.SizeOf<Win32Api.ParentProcessUtilities>(), out _);
                if (current != hProcess) Win32Api.CloseHandle(current);
                if (status != 0 || processInformation.InheritedFromUniqueProcessId == IntPtr.Zero) break;
                current = Win32Api.OpenProcess((uint) Win32Api.ProcessAccessFlags.QueryInformation,
                    false, (uint) processInformation.InheritedFromUniqueProcessId);
                yield return ((uint)processInformation.InheritedFromUniqueProcessId, GetProcessNameFromHandle(current));
            } while (true);
        }
        
        public static void RecoverAllUserProcesses()
        {
            var runningProcesses = Process.GetProcesses();
            var currentSessionId = Process.GetCurrentProcess().SessionId;

            var sameAsThisSession = runningProcesses.Where(p => p.SessionId == currentSessionId);
            foreach (var proc in sameAsThisSession)
            {
                if (BypassProcessList.Contains($"{proc.ProcessName}.exe".ToLowerInvariant())) continue;
                var hProcess = Win32Api.OpenProcess((uint)Win32Api.ProcessAccessFlags.SetInformation, false, (uint) proc.Id);
                ToggleEfficiencyMode(hProcess, false);
                Win32Api.CloseHandle(hProcess);
            }
        }
    }
}