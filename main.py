from re import findall
from subprocess import Popen, PIPE
from colorama import Fore, Back, Style
import winrm
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from datetime import datetime
import pyodbc, os, csv
from plyer import notification

service_name = 'esealservice'

print(f"""------------------------------------------------
        Tracking Server Service Status
------------------------------------------------""")

def check_service_status(ipaddr, username, password, service_name):
    """Check the status of a service on a remote PC."""
    try:
        # Construct WinRM URL
        remote_host = ipaddr
        
        # Create WinRM session
        session = winrm.Session(remote_host, auth=(username, password), transport='ntlm')
        
        # PowerShell script to get service status
        ps_script = f"""
        $service = Get-Service -Name '{service_name}'
        $service.Status
        """
        
        # Execute the PowerShell command
        result = session.run_ps(ps_script)
        
        # Decode and return the result
        return result.std_out.decode().strip()
    except ConnectionError as c:
        return f"Error: {str(c)}"
    except Exception as e:
        return f"Error: {str(e)}"

def get_pc_info(computer_name, username, password):
    """Fetch PC hardware and OS info via WinRM."""
    try:
        session = winrm.Session(computer_name, auth=(username, password), transport='ntlm')
        
        ps_script = """
        $sys = Get-WmiObject Win32_ComputerSystem
        $bios = Get-WmiObject Win32_BIOS
        $cpu = Get-WmiObject Win32_Processor | Select-Object -First 1
        $disks = Get-WmiObject Win32_DiskDrive
        $mem = Get-WmiObject Win32_PhysicalMemory
        $os = Get-WmiObject Win32_OperatingSystem

        $make = $sys.Manufacturer
        $model = $sys.Model
        $hname = $sys.Name
        $serial = $bios.SerialNumber
        $cpu_name = $cpu.Name
        $disk_count = $disks.Count
        $total_disk_size = ($disks | Measure-Object -Property Size -Sum).Sum / 1GB
        $total_mem = ($mem | Measure-Object -Property Capacity -Sum).Sum / 1GB
        $osver = $os.Caption + " " + $os.Version

        "$make|$model|$hname|$serial|$cpu_name|$disk_count|{0:N2}|{1:N2}|$osver" -f $total_disk_size, $total_mem
        """
        
        result = session.run_ps(ps_script)
        output = result.std_out.decode().strip()

        if "|" in output:
            make, model, hname, serial, cpu_name, disk_count, disk_gb, mem_gb, osver = output.split("|")
            return {
                "Make": make,
                "Model":model,
                "Host Name":hname,
                "Serial": serial,
                "Processor": cpu_name,
                "Disk Count": disk_count,
                "Disk Size (GB)": disk_gb,
                "Total Memory (GB)": mem_gb,
                "OS Version": osver
            }
        else:
            return {"error": output}
        
    except Exception as e:
        return {"error": str(e)}
    
def get_disk_details(computer_name, username, password):
    """Fetch detailed disk information from remote PC via WinRM."""
    try:
        session = winrm.Session(computer_name, auth=(username, password), transport='ntlm')
        
        ps_script = """
        $drives = Get-WmiObject Win32_DiskDrive
        $output = @()
        foreach ($drive in $drives) {
            $model = $drive.Model
            $interface = $drive.InterfaceType
            $sizeGB = [math]::Round($drive.Size / 1GB, 2)
            $partitions = $drive.Partitions
            $index = $drive.Index
            $serial = (Get-WmiObject Win32_PhysicalMedia | Where-Object { $_.Tag -eq $drive.DeviceID }).SerialNumber
            
            $output += "$index|$model|$interface|$sizeGB|$partitions|$serial"
        }
        $output -join "`n"
        """
        
        result = session.run_ps(ps_script)
        output = result.std_out.decode().strip()

        if not output:
            return [{"error": "No disk info returned"}]

        disks = []
        for line in output.splitlines():
            fields = line.split("|")
            if len(fields) == 6:
                disks.append({
                    "Index": fields[0],
                    "Model": fields[1].strip(),
                    "Interface": fields[2].strip(),
                    "Size (GB)": fields[3],
                    "Partitions": fields[4],
                    "Serial Number": fields[5].strip()
                })
        return disks
    except Exception as e:
        return [{"error": str(e)}]

def get_resource_utilization(computer_name, username, password):
        """Get CPU load and memory usage stats from the remote machine."""
        try:
            session = winrm.Session(computer_name, auth=(username, password), transport='ntlm')
            
            ps_script = """
            $cpu = (Get-WmiObject Win32_Processor | Measure-Object -Property LoadPercentage -Average).Average
            $mem = Get-WmiObject Win32_OperatingSystem
            $total = [math]::Round($mem.TotalVisibleMemorySize / 1024, 2)
            $free = [math]::Round($mem.FreePhysicalMemory / 1024, 2)
            $used = [math]::Round($total - $free, 2)
            $mem_percent = [math]::Round(($used / $total) * 100, 2)
            "$cpu|$used|$free|$total|$mem_percent"
            """
            result = session.run_ps(ps_script)
            output = result.std_out.decode().strip()

            if "|" in output:
                cpu, used, free, total, mem_percent = output.split("|")
                return {
                    "cpu_percent": cpu,
                    "mem_used_mb": used,
                    "mem_free_mb": free,
                    "mem_total_mb": total,
                    "mem_percent": mem_percent
                }
            else:
                return {"error": output}
        except Exception as e:
            return {"error": str(e)}

def show_popup(title, message):
    # Display a pop-up alert on Windows.
    notification.notify(
        title=title,
        message=message,
        app_name='Service Monitor',
        timeout=5  # seconds
    )

def write_log(ipaddr, hostname, loc, plcod, pc_sts, rslt, timestamp, status, CPU, memory):
    
    log_dir = "C:\\Log Files"
    csv_log_dir = "C:\\Log Files\\csv_log"
    csv_file = os.path.join(csv_log_dir, f"service_status_log_{datetime.now().strftime('%Y-%m-%d')}.csv")
    log_file = os.path.join(log_dir, f"service_status_{datetime.now().strftime('%Y-%m-%d')}.log")

    # Ensure the log directory exists
    os.makedirs(log_dir, exist_ok=True)
    
    # Check if file exists to write headers
    file_exists = os.path.isfile(csv_file)

    with open(log_file, "a", encoding="utf-8") as f:
        f.write(f"[{timestamp}] IP: {ipaddr} | Hostname: {hostname} | Location: {loc} | PlantCode: {plcod} | PC Status: {pc_sts} | Service Status: {rslt} | CPU Utilize: {CPU} | Memory Utilize: {memory} | Other_Errors: {status}\n")
        
    with open(csv_file, mode='a', newline='', encoding='utf-8') as f:
        writer = csv.writer(f)
        
        # Write header if the file is new
        if not file_exists:
            writer.writerow(["Timestamp", "IP Address", "Hostname", "Location", "Plant Code", "PC Status", "Service Status", "CPU Utilize", "Memory Utilize","Other_Errors"])
            
            # Write the log data
        writer.writerow([timestamp, ipaddr, hostname, loc, plcod, pc_sts, rslt, CPU, memory, status])

def send_email(subject, body, to_email):
    """Send an email with the specified subject and body."""
    from_email = "any@mail.com"  # Replace with your email
    from_password = "password123:"  # Replace with your email password

    # Create the email message
    msg = MIMEMultipart()
    msg['From'] = from_email
    msg['To'] = ", ".join(to_email)
    msg['Subject'] = subject

    # Attach the email body
    msg.attach(MIMEText(body, 'plain'))

    # Send the email
    try:
        server = smtplib.SMTP('smtp.office365.com', 25)  # Replace with your SMTP server and port
        server.starttls()
        server.login(from_email, from_password)
        text = msg.as_string()
        server.sendmail(from_email, to_email, text)
        server.quit()
        print(f"Email sent to {to_email}")
    except Exception as e:
        print(f"Failed to send email: {str(e)}")

while True:
    try:
        conn = pyodbc.connect(
            'DRIVER={ODBC Driver 17 for SQL Server};'
            'SERVER=hostnameoripaddress\\SQLEXPRESS;'
            'DATABASE=dbname;'
            'UID=sa;'
            'PWD=password123;',  # Replace with your SQL Server credentials
            timeout=5  # optional: connection timeout in seconds
        )
        
        cursor = conn.cursor()
        # Execute a SELECT query
        cursor.execute("SELECT DesktopIPAddrs, Hostname, LocationName, PlantCode, AdminPasswd FROM DSPCdetails")  # change table/columns as needed
        # Fetch all rows
        rows = cursor.fetchall()
        
        for row in rows:
            ipaddr = row.DesktopIPAddrs
            hostname = row.Hostname
            loc = row.LocationName
            plcod = row.PlantCode
            username = "administrator"
            password = row.AdminPasswd
            to_email = "mail@domain.com"  # Add this column to your CSV
            timestamp = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
            
            data = ""
            output = Popen(f"ping {ipaddr} -n 5", stdout=PIPE, encoding="utf-8")
            for line in output.stdout:
                data = data + line
                ping_test = findall("TTL", data)
            if ping_test:
                pc_status = f"{Fore.GREEN}Active{Style.RESET_ALL}"
                pc_sts = "Active"
            else:
                pc_status = f"{Fore.RED}Down  {Style.RESET_ALL}"
                pc_sts = "Down"
                
                # Call the popup alert
                show_popup(
                    f"System is Down on {hostname}:{ipaddr}",
                    f"{service_name} is Down at {loc} ({plcod})"
                )
            res_util = get_resource_utilization(ipaddr, username, password)
            
            if "error" not in res_util:
                cpu_usage = res_util["cpu_percent"]
                mem_usage = res_util["mem_percent"]
                total_mem = res_util["mem_total_mb"]
                if mem_usage >= '90':
                    mem = f"{Fore.RED}{mem_usage}%{Style.RESET_ALL}"
                    memory = f"{mem_usage}%"
                    tot_mem = f"{Fore.CYAN}{total_mem}{Style.RESET_ALL}"
                    CPU = f"{cpu_usage}%"
                    cpu = f"{cpu_usage}% "
                    show_popup(
                        f"Memory Utilize on {loc} - {ipaddr}",
                        f"Memry Utilization is high: {memory}"
                    )
                elif mem_usage >= '70':
                    mem = f"{Fore.YELLOW}{mem_usage}%{Style.RESET_ALL}"
                    memory = f"{mem_usage}%"
                    tot_mem = f"{Fore.CYAN}{total_mem}{Style.RESET_ALL}"
                    CPU = f"{cpu_usage}%"
                    cpu = f"{cpu_usage}% "
                else:
                    cpu = f"{Fore.GREEN}{cpu_usage}% {Style.RESET_ALL}"
                    mem = f"{Fore.GREEN}{mem_usage}%{Style.RESET_ALL}"
                    memory = f"{mem_usage}%"
                    tot_mem = f"{Fore.CYAN}{total_mem}{Style.RESET_ALL}"
                    CPU = f"{cpu_usage}%"
            else:
                cpu = f"{Fore.YELLOW}Error {Style.RESET_ALL}"
                mem = f"{Fore.YELLOW}Error {Style.RESET_ALL}"
                tot_mem = f"{Fore.YELLOW}Error {Style.RESET_ALL}"
                CPU = "Resource_Fetch_Error"
                memory = "Resource_Fetch_Error"
            status = check_service_status(ipaddr, username, password, service_name)
            if status == "Running":
                result = f"{Fore.GREEN}Running{Style.RESET_ALL}"
                rslt = "Running"
            elif status == "Stopped":
                result = f"{Fore.RED}Stopped{Style.RESET_ALL} {status}"
                rslt = "Stopped"
                
                # Call the popup alert
                show_popup(
                    f"Service Stopped on {hostname}:{ipaddr}",
                    f"{service_name} is stopped at {loc} ({plcod})"
                )

                # Prepare email subject and body
                subject = f"Service Status for {service_name} on {ipaddr}"
                body = f"""
                Timestamp: {timestamp}
                IP Address: {ipaddr}
                Hostname: {hostname}
                Location: {loc}
                Plant Code: {plcod}
                PC Status: {pc_sts}
                Service Status: {rslt}
                """
                # Send email
                send_email(subject, body, to_email)
            else:
                result = f"{Fore.RED}Down   {Style.RESET_ALL}"
                rslt = "Down"
                
                # Call the popup alert
                show_popup(
                    f"Service Down on {loc} - {ipaddr}",
                    "Other_Error: Winrm Communication or Time-out error"
                )
            write_log(ipaddr, hostname, loc, plcod, pc_sts, rslt, timestamp, status, CPU, memory)
            print(f"""------------------------------------------------
TimeStamp  : {Fore.MAGENTA}{timestamp}{Style.RESET_ALL}
Host Name  : {Fore.CYAN}{hostname}{Style.RESET_ALL}
IP Address : {Fore.CYAN}{ipaddr} {Style.RESET_ALL}
Location   : {Fore.CYAN}{loc}{Style.RESET_ALL}
PlantCode  : {Fore.CYAN}{plcod}{Style.RESET_ALL}
PC Status  : {pc_status}  | Service    : {result}
TotalMemory: {tot_mem} | Memory Util: {mem}
------------------------------------------------""")
            
            # Cleanup
        cursor.close()
        conn.close()  # optional: close immediately if just testing
        
    except pyodbc.Error as e:
        print("Connection failed.")
        print("SQL_Error:", e)          