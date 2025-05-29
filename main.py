import serial
import time
from openpyxl import Workbook
from datetime import datetime
import tkinter as tk
from tkinter import messagebox
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
import matplotlib.pyplot as plt
import threading
import serial.tools.list_ports

def setup_arduino():
    try:
        # List all available ports
        ports = list(serial.tools.list_ports.comports())
        print("Available ports:")
        for port in ports:
            print(f"- {port.device}: {port.description}")
        
        # Try to connect to COM4
        ser = serial.Serial('COM4', 9600, timeout=1)
        time.sleep(2)  # Wait for Arduino to reset
        print("Successfully connected to Arduino!")
        return ser
    except serial.SerialException as e:
        print(f"Error connecting to Arduino: {e}")
        print("\nPlease check:")
        print("1. Is Arduino IDE's Serial Monitor closed?")
        print("2. Is the correct port selected?")
        print("3. Is Arduino powered on?")
        return None

def parse_sensor_data(data):
    voltage = current = tds = temp = error = ""
    try:
        print(f"Raw data received: {data}")
        # Split the data by spaces and process each part
        parts = data.split()
        for i, part in enumerate(parts):
            if "$Voltage$" in part:
                # Extract voltage value and remove 'V'
                voltage = parts[i+2].replace("V", "").strip()
                print(f"Parsed voltage: {voltage}")
            elif "$TDS$" in part:
                # Extract TDS value
                tds = parts[i+2].strip()
                print(f"Parsed TDS: {tds}")
            elif "$Temp$" in part:
                temp = parts[i+2].strip()
                print(f"Parsed temperature: {temp}")
    except Exception as e:
        print(f"Error parsing data: {e}")
    return voltage, current, tds, temp, error

def start_logging(ax, canvas, wb, ws, filename, ser, stop_event, root):
    s_no = 1
    voltages = []
    currents = []
    timelapses = []
    start_time = time.time()
    print(f"Starting logging to file: {filename}")
    
    # Initialize the plot
    ax.set_xlabel('Timelapse (s)')
    ax.set_ylabel('Value')
    ax.set_title('Live Voltage vs Timelapse')
    ax.grid(True)
    canvas.draw()
    
    try:
        while not stop_event.is_set():
            if ser.in_waiting:
                try:
                    data = ser.readline().decode('utf-8').strip()
                    if data:
                        print(f"Received data: {data}")
                        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                        voltage, current, tds, temp, error = parse_sensor_data(data)
                        
                        # Process data if we have valid voltage or TDS
                        if voltage or tds:
                            try:
                                # Convert values to float if they exist
                                v = float(voltage) if voltage else ""
                                tds_val = float(tds) if tds else ""
                                t = time.time() - start_time
                                
                                print(f"Processing values - Voltage: {v}, TDS: {tds_val}")
                                
                                # Append data to lists if we have voltage
                                if voltage:
                                    voltages.append(v)
                                    timelapses.append(t)
                                
                                # Write to Excel
                                row_data = [s_no, timestamp, v, current, tds_val, temp, error]
                                print(f"Writing to Excel: {row_data}")
                                ws.append(row_data)
                                wb.save(filename)
                                print(f"Data saved to Excel, row {s_no}")
                                s_no += 1
                                
                                # Update plot if we have voltage data
                                if voltage and len(voltages) > 0:
                                    ax.clear()
                                    ax.plot(timelapses, voltages, label='Voltage (V)', color='blue', linewidth=2)
                                    ax.set_xlabel('Timelapse (s)')
                                    ax.set_ylabel('Voltage (V)')
                                    ax.set_title('Live Voltage vs Timelapse')
                                    ax.grid(True)
                                    ax.legend()
                                    
                                    # Set y-axis limits with some padding
                                    min_val = min(voltages)
                                    max_val = max(voltages)
                                    padding = (max_val - min_val) * 0.1 if max_val != min_val else 0.1
                                    ax.set_ylim(min_val - padding, max_val + padding)
                                    
                                    # Set x-axis limits
                                    ax.set_xlim(min(timelapses), max(timelapses))
                                    
                                    canvas.draw()
                                    root.update_idletasks()
                                    print("Graph updated")
                                
                            except ValueError as e:
                                print(f"Error converting values: {e}")
                            except Exception as e:
                                print(f"Error processing data: {e}")
                        else:
                            print("No valid voltage or TDS data received")
                except serial.SerialException as e:
                    print(f"Serial port error: {e}")
                    break
                except Exception as e:
                    print(f"Error reading serial data: {e}")
                
                time.sleep(0.1)  # Small delay to prevent overwhelming the system
    except Exception as e:
        print(f"Error in logging loop: {e}")
        if not stop_event.is_set():  # Only show error if not stopping
            try:
                messagebox.showerror("Error", str(e))
            except:
                print("Could not show error message box")
    finally:
        ser.close()
        if not stop_event.is_set():  # Only show message if not stopping
            try:
                messagebox.showinfo("Stopped", "Serial connection closed.")
            except:
                print("Could not show message box")

def main():
    # Setup Arduino
    ser = setup_arduino()
    if not ser:
        return

    # Setup Excel
    wb = Workbook()
    ws = wb.active
    ws.title = "SensorData"
    headers = ["S.no", "Timestamp", "Voltage", "Current", "TDS", "Temp", "Error"]
    ws.append(headers)
    
    # Create filename with timestamp
    filename = f"sensor_data_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
    print(f"Created Excel file: {filename}")
    wb.save(filename)
    
    # Tkinter GUI
    root = tk.Tk()
    root.title("Arduino Excel Logger - Live Graph")
    root.geometry("800x600")  # Made window slightly larger

    label = tk.Label(root, text="Live Voltage vs Timelapse", font=("Arial", 14))
    label.pack(pady=5)

    # Create figure with larger size
    fig, ax = plt.subplots(figsize=(8, 4))
    canvas = FigureCanvasTkAgg(fig, master=root)
    canvas.get_tk_widget().pack(padx=10, pady=5, fill=tk.BOTH, expand=True)

    stop_event = threading.Event()
    logging_thread = None

    def on_start():
        nonlocal logging_thread
        print("Start button clicked")
        start_button.config(state=tk.DISABLED)
        logging_thread = threading.Thread(target=start_logging, args=(ax, canvas, wb, ws, filename, ser, stop_event, root))
        logging_thread.daemon = True
        logging_thread.start()
        print("Logging thread started")

    def on_close():
        print("Window closing")
        stop_event.set()
        root.destroy()

    start_button = tk.Button(root, text="Start", font=("Arial", 12), width=10, command=on_start)
    start_button.pack(pady=10)

    root.protocol("WM_DELETE_WINDOW", on_close)
    root.mainloop()

if __name__ == "__main__":
    main()
