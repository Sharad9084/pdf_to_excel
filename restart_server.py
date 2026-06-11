import subprocess
import os
import sys
import time

print("=== RESTARTING DATABASE SERVER ===")

# 1. Find PIDs on port 8787
pids = set()
try:
    output = subprocess.check_output("netstat -ano | findstr 8787", shell=True).decode('utf-8')
    for line in output.strip().split('\n'):
        parts = line.split()
        if len(parts) >= 5 and "LISTENING" in parts[3]:
            pid = parts[-1]
            pids.add(int(pid))
except Exception as e:
    print(f"No active process found on port 8787 or error: {e}")

print(f"PIDs to kill: {pids}")

# 2. Kill the processes
for pid in pids:
    try:
        print(f"Killing PID {pid}...")
        subprocess.check_call(f"taskkill /F /PID {pid}", shell=True)
        print(f"Successfully killed PID {pid}")
    except Exception as e:
        print(f"Failed to kill PID {pid}: {e}")

time.sleep(2)

# 3. Start the correct server.py in backend
server_file = r"C:\Users\hp\OneDrive\Desktop\fronted\backend\server.py"
server_dir = r"C:\Users\hp\OneDrive\Desktop\fronted\backend"

print(f"Starting server: {server_file} in {server_dir}")
try:
    # Use subprocess.Popen to start it as a detached background process
    p = subprocess.Popen(
        [sys.executable, "server.py"],
        cwd=server_dir,
        stdout=open(os.path.join(server_dir, "server.out.log"), "w"),
        stderr=open(os.path.join(server_dir, "server.err.log"), "w"),
        creationflags=subprocess.CREATE_NEW_PROCESS_GROUP | subprocess.DETACHED_PROCESS if os.name == 'nt' else 0
    )
    print(f"Started server process with PID {p.pid}")
except Exception as e:
    print(f"Error starting server: {e}")
