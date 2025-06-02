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

def start_logging(ax1, ax2, canvas, wb, ws, filename, ser, stop_event, root, selected_param1, selected_param2, split_screen_var):
    s_no = 1
    voltages = []
    currents = []
    tds_values = []
    temp_values = []
    timelapses = []
    start_time = time.time()
    print(f"Starting logging to file: {filename}")
    
    # Initialize the plots
    ax1.set_xlabel('Timelapse (s)')
    ax1.set_ylabel('Value')
    ax1.set_title(f'Live {selected_param1.get()} vs Timelapse')
    ax1.grid(True)
    
    if split_screen_var.get():
        ax2.set_xlabel('Timelapse (s)')
        ax2.set_ylabel('Value')
        ax2.set_title(f'Live {selected_param2.get()} vs Timelapse')
        ax2.grid(True)
    
    canvas.draw()
    
    try:
        while not stop_event.is_set():
            if ser.in_waiting:
                data = None
                try:
                    data = ser.readline().decode('utf-8').strip()
                except serial.SerialException as e:
                    print(f"Serial port error: {e}")
                    break
                except Exception as e:
                    print(f"Error reading serial data: {e}")
                if data:
                    print(f"Received data: {data}")
                    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                    voltage, current, tds, temp, error = parse_sensor_data(data)
                    # Process data if we have valid values
                    if voltage or tds or temp:
                        try:
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
                            
                            # Update first plot
                            ax1.clear()
                            param1 = selected_param1.get()
                            update_plot(ax1, param1, voltages, tds_values, temp_values, timelapses)
                            
                            # Update second plot if split screen is enabled
                            if split_screen_var.get():
                                ax2.clear()
                                param2 = selected_param2.get()
                                update_plot(ax2, param2, voltages, tds_values, temp_values, timelapses)
                            
                            canvas.draw()
                            root.update_idletasks()
                            print("Graphs updated")
                        except ValueError as e:
                            print(f"Error converting values: {e}")
                        except Exception as e:
                            print(f"Error processing data: {e}")
                    else:
                        print("No valid data received")
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

def update_plot(ax, param, voltages, tds_values, temp_values, timelapses):
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
        plot_times = timelapses[:len(data_values)]
        ax.plot(plot_times, data_values, label=label, color=color, linewidth=2)
        ax.set_ylabel(ylabel)
        ax.set_xlabel('Timelapse (s)')
        ax.set_title(f'Live {param} vs Timelapse')
        ax.grid(True)
        ax.legend()
        # Set y-axis limits with some padding
        min_val = min(data_values)
        max_val = max(data_values)
        padding = (max_val - min_val) * 0.1 if max_val != min_val else 0.1
        ax.set_ylim(min_val - padding, max_val + padding)
        # Set x-axis limits
        ax.set_xlim(min(plot_times), max(plot_times))

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
    root.geometry("1000x800")  # Made window larger for split screen

    # Create a frame for the controls
    control_frame = tk.Frame(root)
    control_frame.pack(pady=5)

    label = tk.Label(control_frame, text="Live Data vs Timelapse", font=("Arial", 14))
    label.pack(side=tk.LEFT, padx=10)

    # Create dropdowns for parameter selection
    selected_param1 = tk.StringVar(value="Voltage")
    selected_param2 = tk.StringVar(value="TDS")
    
    param_label1 = tk.Label(control_frame, text="Graph 1:", font=("Arial", 10))
    param_label1.pack(side=tk.LEFT, padx=5)
    param_dropdown1 = tk.OptionMenu(control_frame, selected_param1, "Voltage", "TDS", "Temperature")
    param_dropdown1.pack(side=tk.LEFT, padx=5)
    
    # Split screen checkbox
    split_screen_var = tk.BooleanVar(value=False)
    split_screen_check = tk.Checkbutton(control_frame, text="Split Screen", variable=split_screen_var, 
                                      command=lambda: update_split_screen(split_screen_var, param_label2, param_dropdown2))
    split_screen_check.pack(side=tk.LEFT, padx=5)
    
    # Second parameter dropdown (initially hidden)
    param_label2 = tk.Label(control_frame, text="Graph 2:", font=("Arial", 10))
    param_dropdown2 = tk.OptionMenu(control_frame, selected_param2, "TDS", "Temperature")
    param_label2.pack_forget()
    param_dropdown2.pack_forget()

    # Create figure with subplots
    fig, (ax1, ax2) = plt.subplots(2, 1, figsize=(10, 8))
    canvas = FigureCanvasTkAgg(fig, master=root)
    canvas.get_tk_widget().pack(padx=10, pady=5, fill=tk.BOTH, expand=True)

    stop_event = threading.Event()
    logging_thread = None

    def update_split_screen(split_var, label2, dropdown2):
        if split_var.get():
            label2.pack(side=tk.LEFT, padx=5)
            dropdown2.pack(side=tk.LEFT, padx=5)
            ax2.set_visible(True)
        else:
            label2.pack_forget()
            dropdown2.pack_forget()
            ax2.set_visible(False)
        canvas.draw()

    def on_start():
        nonlocal logging_thread
        print("Start button clicked")
        start_button.config(state=tk.DISABLED)
        logging_thread = threading.Thread(target=start_logging, 
                                        args=(ax1, ax2, canvas, wb, ws, filename, ser, stop_event, 
                                             root, selected_param1, selected_param2, split_screen_var))
        logging_thread.daemon = True
        logging_thread.start()
        print("Logging thread started")

    # Add Start button to control frame
    start_button = tk.Button(control_frame, text="Start Logging", font=("Arial", 12, "bold"), 
                           width=15, command=on_start, bg="#4CAF50", fg="white")
    start_button.pack(side=tk.LEFT, padx=20)

    def on_close():
        print("Window closing")
        stop_event.set()
        root.destroy()

    root.protocol("WM_DELETE_WINDOW", on_close)
    root.mainloop()

if __name__ == "__main__":
    main()
