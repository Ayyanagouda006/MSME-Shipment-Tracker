import time
import subprocess
 
def run_script():
    # Replace with the full path to your script
	script_path = "/MSME-Shipment-Tracker/Backend_data.py"
    
    try:
		subprocess.run(["python3", script_path], check=True)
        print(f"Executed {script_path} successfully.")
    except subprocess.CalledProcessError as e:
        print(f"Error executing script: {e}")
 
if __name__ == "__main__":
    while True:
        run_script()
        time.sleep(3600) 