import subprocess
import os

print("=== CHECKING PORT 8787 ===")
try:
    # Use netstat to find port 8787
    output = subprocess.check_output("netstat -ano | findstr 8787", shell=True).decode('utf-8')
    print("Netstat output:")
    print(output)
    
    # Extract PID
    pids = set()
    for line in output.strip().split('\n'):
        parts = line.split()
        if len(parts) >= 5:
            pid = parts[-1]
            pids.add(pid)
            
    print(f"PIDs listening on 8787: {pids}")
    for pid in pids:
        try:
            task_info = subprocess.check_output(f"tasklist /FI \"PID eq {pid}\"", shell=True).decode('utf-8')
            print(f"\nTask info for PID {pid}:")
            print(task_info)
            
            # Get command line for process if possible
            cmd_info = subprocess.check_output(f"wmic process where processid={pid} get commandline", shell=True).decode('utf-8')
            print(f"Command line for PID {pid}:")
            print(cmd_info.strip())
        except Exception as pe:
            print(f"Error checking PID {pid}: {pe}")
except Exception as e:
    print(f"Error checking port: {e}")
