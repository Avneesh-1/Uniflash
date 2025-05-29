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
                tds = parts[i+2].strip()
                print(f"Parsed TDS: {tds}")
            elif "$Temp$" in part:
                temp = parts[i+2].strip()
                print(f"Parsed temperature: {temp}")
    except Exception as e:
        print(f"Error parsing data: {e}")
    return voltage, current, tds, temp, error

def start_logging(ax, canvas, wb, ws, filename, ser, stop_event, root, selected_param):
    s_no = 1
    voltages = []
    currents = []
    tds_values = []
    temp_values = []
    timelapses = []
    start_time = time.time()
    print(f"Starting logging to file: {filename}")
    
    # Initialize the plot
    ax.set_xlabel('Timelapse (s)')
    ax.set_ylabel('Value')
    ax.set_title(f'Live {selected_param.get()} vs Timelapse')
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
                        
                        # Process data if we have valid values
                        if voltage or tds or temp:
                            try:
                                # Convert values to float if they exist
                                v = float(voltage) if voltage else None
                                tds_val = float(tds) if tds else None
                                temp_val = float(temp) if temp else None
                                t = time.time() - start_time
                                
                                print(f"Processing values - Voltage: {v}, TDS: {tds_val}, Temp: {temp_val}")
                                
                                # Append data to lists
                                if v is not None:
                                    voltages.append(v)
                                    timelapses.append(t)
                                if tds_val is not None:
                                    tds_values.append(tds_val)
                                    if len(tds_values) > len(timelapses):
                                        timelapses.append(t)
                                if temp_val is not None:
                                    temp_values.append(temp_val)
                                    if len(temp_values) > len(timelapses):
                                        timelapses.append(t)
                                
                                # Write to Excel
                                row_data = [s_no, timestamp, v, current, tds_val, temp_val, error]
                                print(f"Writing to Excel: {row_data}")
                                ws.append(row_data)
                                wb.save(filename)
                                print(f"Data saved to Excel, row {s_no}")
                                s_no += 1
                                
                                # Update plot based on selected parameter
                                ax.clear()
                                param = selected_param.get()
                                
                                # Get the appropriate data arrays based on selected parameter
                                if param == "Voltage" and len(voltages) > 0:
                                    data_values = voltages
                                    label = 'Voltage (V)'
                                    color = 'blue'
                                    ylabel = 'Voltage (V)'
                                elif param == "TDS" and len(tds_values) > 0:
                                    data_values = tds_values
                                    label = 'TDS'
                                    color = 'red'
                                    ylabel = 'TDS'
                                elif param == "Temperature" and len(temp_values) > 0:
                                    data_values = temp_values
                                    label = 'Temperature (°C)'
                                    color = 'green'
                                    ylabel = 'Temperature (°C)'
                                else:
                                    data_values = []
                                
                                if data_values:
                                    # Ensure we have matching lengths
                                    plot_times = timelapses[:len(data_values)]
                                    ax.plot(plot_times, data_values, label=label, color=color, linewidth=2)
                                    ax.set_ylabel(ylabel)
                                
                                ax.set_xlabel('Timelapse (s)')
                                ax.set_title(f'Live {param} vs Timelapse')
                                ax.grid(True)
                                ax.legend()
                                
                                # Set y-axis limits with some padding
                                if data_values:
                                    min_val = min(data_values)
                                    max_val = max(data_values)
                                    padding = (max_val - min_val) * 0.1 if max_val != min_val else 0.1
                                    ax.set_ylim(min_val - padding, max_val + padding)
                                    
                                    # Set x-axis limits
                                    ax.set_xlim(min(plot_times), max(plot_times))
                                
                                canvas.draw()
                                root.update_idletasks()
                                print("Graph updated")
                                
                            except ValueError as e:
                                print(f"Error converting values: {e}")
                            except Exception as e:
                                print(f"Error processing data: {e}")
                        else:
                            print("No valid data received")
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

    # Create a frame for the controls
    control_frame = tk.Frame(root)
    control_frame.pack(pady=5)

    label = tk.Label(control_frame, text="Live Data vs Timelapse", font=("Arial", 14))
    label.pack(side=tk.LEFT, padx=10)

    # Create dropdown for parameter selection
    selected_param = tk.StringVar(value="Voltage")
    param_label = tk.Label(control_frame, text="Select Parameter:", font=("Arial", 10))
    param_label.pack(side=tk.LEFT, padx=5)
    param_dropdown = tk.OptionMenu(control_frame, selected_param, "Voltage", "TDS", "Temperature")
    param_dropdown.pack(side=tk.LEFT, padx=5)

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
        logging_thread = threading.Thread(target=start_logging, args=(ax, canvas, wb, ws, filename, ser, stop_event, root, selected_param))
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
