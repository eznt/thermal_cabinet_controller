# 14.7.2023
import re
import tkinter as tk
import pyvisa
import time
import threading
import datetime
import csv

from tkinter import ttk
from tkinter import messagebox


class SU241:
    def __init__(self, gbip_addr: str):
        self.gbip_addr = gbip_addr
        self.resource = None
        self.rm = pyvisa.ResourceManager()
        self.resources = self.rm.list_resources()
        self.instrument = None
        self.instrument_connected = False
        self.busy = False
        self.program_running = False
        self.mode = None
        self.creating_program = False
        self.lock = threading.Lock()

    def connect(self):
        try:
            self.resource = pyvisa.ResourceManager().open_resource(self.gbip_addr, timeout=5000)
            self.instrument_connected = True
            self.mode = self.query("MODE?").replace("\r\n", "")
            if "RUN" in self.mode:
                self.program_running = True
        except pyvisa.Error as e:
            print(f"An error occured while connecting {e}")
            self.instrument_connected = False
            return

    def send_cmd(self, cmd):
        if self.instrument_connected:
            self.lock.acquire()
            self.resource.write(cmd)
            time.sleep(0.5)
            r = self.resource.read()
            print(f'{cmd}: {r}')
            if r.startswith('OK') and "ERR-" not in r:
                self.lock.release()
                return 'OK'
            self.lock.release()
            return 'NOK'
        else:
            print("Not connected to instrument")

    def query(self, q):
        if self.instrument_connected:
            self.lock.acquire()
            self.resource.write(q)
            time.sleep(0.5)
            r = self.resource.read()
            self.lock.release()
            return r if not r.startswith('NOK') else 'NOK'
        else:
            print("Not connected to instrument")

    def id(self):
        return self.query('TYPE?')

    def set_mode(self, m):
        return self.send_cmd(f'MODE, {m}')

    def power_on(self):
        return self.send_cmd('POWER, ON')

    def power_off(self):
        return self.send_cmd('POWER, OFF')

    def set_temp(self, set_t, high_t=85, low_t=-45):
        return self.send_cmd(f'TEMP, S{set_t:.1f} H{high_t:.1f} L{low_t:.1f}')

    def read_temp(self):
        r = self.query('TEMP?')
        if r != 'NOK':
            temp_values = list(map(float, r.split(',')))
            return {'current': temp_values[0], 'set': temp_values[1], 'high': temp_values[2], 'low': temp_values[3]}
        else:
            return None

    def get_program_status(self):
        result = self.query("PRGM MON?")
        if "NA:CONTROLLER NOT READY-2" in result:
            print("Program not running")
            parsed_results = {"PROGRAM_NUMBER": 0, "CURRENT_STEP": 0, "TARGET_TEMP": 0,
                              "STEP_TIME_REMAINING": 0, "REPEAT_CYCLE_COUNT": 0}
            return parsed_results
        results = result.split(",")
        parsed_results = {"PROGRAM_NUMBER": results[0], "CURRENT_STEP": results[1], "TARGET_TEMP": results[2],
                          "STEP_TIME_REMAINING": results[3], "REPEAT_CYCLE_COUNT": results[4]}
        return parsed_results

    def write_program(self, program_number=1):
        if self.instrument_connected:
            self.creating_program = True
            response = ""

            response += self.send_cmd(f"PRGM DATA WRITE, PGM: {program_number}, EDIT START")
            if "NOK" in response: self.creating_program = False; return False
            response += self.send_cmd(f"PRGM DATA WRITE, PGM: {program_number}, STEP1, TEMP60.0, TRAMP ON, TIME01:30")
            response += self.send_cmd(f"PRGM DATA WRITE, PGM: {program_number}, STEP2, TIME01:30")
            if "NOK" in response: self.creating_program = False; return False
            response += self.send_cmd(f"PRGM DATA WRITE, PGM: {program_number}, STEP3, TEMP-40.0, TRAMP ON, TIME01:30")
            response += self.send_cmd(f"PRGM DATA WRITE, PGM: {program_number}, STEP4, TIME01:30")
            if "NOK" in response: self.creating_program = False; return False
            response += self.send_cmd(f"PRGM DATA WRITE, PGM: {program_number}, STEP5, TEMP23.0, TRAMP ON, TIME01:30")
            response += self.send_cmd(f"PRGM DATA WRITE, PGM: {program_number}, COUNT, (1. 4. 2)")
            if "NOK" in response: self.creating_program = False; return False
            response += self.send_cmd(f"PRGM DATA WRITE, PGM: {program_number}, END, HOLD")
            response += self.send_cmd(f"PRGM DATA WRITE, PGM: {program_number}, EDIT END")
            if "NOK" in response: self.creating_program = False; return False

            """
            self.send_cmd(f"PRGM DATA WRITE, PGM: {program_number}, EDIT START")
            self.send_cmd(f"PRGM DATA WRITE, PGM: {program_number}, STEP1, TEMP18.0, TRAMP ON, TIME00:02")
            self.send_cmd(f"PRGM DATA WRITE, PGM: {program_number}, STEP2, TIME00:01")
            self.send_cmd(f"PRGM DATA WRITE, PGM: {program_number}, STEP3, TEMP20.0, TRAMP ON, TIME00:02")
            self.send_cmd(f"PRGM DATA WRITE, PGM: {program_number}, STEP4, TIME00:01")
            self.send_cmd(f"PRGM DATA WRITE, PGM: {program_number}, END, HOLD")
            self.send_cmd(f"PRGM DATA WRITE, PGM: {program_number}, EDIT END")
            """
            self.creating_program = False
            return True
        else:
            print("Not connected to instrument")

    def run_program(self, program_number):
        output = self.send_cmd(f'MODE, RUN {program_number}')
        if output is not None:
            print(f"Program {program_number} started")
            self.program_running = True

    def create_new_program(self, steps):
        pass

    def rewrite_program(self, steps):
        for step in steps:
            self.send_cmd(step)
            time.sleep(0.5)

    PROGRAM_STEPS = [
        "PRGM DATA WRITE, PGM: 1, EDIT START",
        "PRGM DATA WRITE, PGM: 1, STEP1, TEMP40.0, TRAMP ON, TIME00:15",
        "PRGM DATA WRITE, PGM: 1, END, HOLD",
        "PRGM DATA WRITE, PGM: 1, EDIT END",
    ]

    def is_busy(self):
        return self.busy

    def is_connected(self):
        return self.instrument_connected

    def is_program_running(self):
        return self.program_running

    def change_mode(self, mode):
        self.mode = mode

    def get_mode(self):
        return self.mode


class MainWindow:
    def __init__(self, cabinet: SU241):
        self.window = tk.Tk()
        self.window.title("Temp Control")
        self.window.geometry("640x320")
        self.cabinet = cabinet
        self.output_thread = None
        self.is_closed = False
        self.window.protocol("WM_DELETE_WINDOW", self.close)

        self.lock = threading.Lock()

        self.validation = self.window.register(self.validate_numeric_input)

        self.instrument_label = tk.Label(self.window, text="Instrument: SU-241", font=("Helvetica", 14))
        self.instrument_label.place(x=10, y=20)

        self.connection_label = tk.Label(self.window, text="No Connection", font=("Helvetica", 14), bg="white",
                                         width=40)
        self.connection_label.place(x=160, y=73)

        self.output_text = tk.Text(self.window, height=10, width=50)
        self.output_text.place(x=10, y=120)

        self.input_entry = ttk.Entry(self.window, font=("Helvetica", 14), validate="key",
                                     validatecommand=(self.validation, '%S'))
        self.input_entry.place(x=180, y=250)

        self.connect_button = tk.Button(self.window, text="Connect Cabinet", width=14, bd=4, font=("Helvetica", 11),
                                        command=self.connect_cabinet)
        self.connect_button.place(x=10, y=70)

        self.set_temperature_button = tk.Button(self.window, text="Set Constant Temp", width=14, bd=4,
                                                font=("Helvetica", 11),
                                                command=self.set_constant_temperature)
        self.set_temperature_button.place(x=440, y=170)

        self.rewrite_program_button = tk.Button(self.window, text="Create Ramp", width=14, bd=4, font=("Helvetica", 11),
                                                command=self.rewrite_program_steps)
        self.rewrite_program_button.place(x=440, y=130)

        self.run_program_button = tk.Button(self.window, text="Start Ramp", width=14, bd=4, font=("Helvetica", 11),
                                            command=self.run_program)

        self.stop_program_button = tk.Button(self.window, text="Stop Ramp", width=14, bd=4, font=("Helvetica", 11),
                                             command=self.stop_program)

        self.set_output_on_button = tk.Button(self.window, text="Standby", width=14, bd=4, font=("Helvetica", 11),
                                              command=lambda: self.power_on())
        self.set_output_on_button.place(x=440, y=212)

        self.set_output_off_button = tk.Button(self.window, text="Off", width=14, bd=4, font=("Helvetica", 11),
                                               command=lambda: self.power_off(), bg="red")
        self.set_output_off_button.place(x=440, y=252)

        self.window.mainloop()

    def validate_numeric_input(self, char):
        if char.isdigit():
            return True
        else:
            return False

    def connect_cabinet(self):
        self.cabinet.connect()
        time.sleep(1)
        self.connection_label.config(text="Connected")
        if self.cabinet.is_connected():
            mode = self.cabinet.get_mode()
            self.switch_activated_buttons(mode)
            self.start_update_output_thread()
            self.connect_button.config(state=tk.DISABLED)
            if self.cabinet.program_running:
                self.stop_program_button.place(x=440, y=100)
            else:
                self.run_program_button.place(x=440, y=100)

    def set_constant_temperature(self):
        self.lock.acquire()
        temperature_str = self.input_entry.get()
        try:
            temperature = float(temperature_str)
        except ValueError as e:
            self.input_entry.insert(0, "Input a valid temperature")
            print("Input a valid temperature")
            self.lock.release()
            return
        self.cabinet.set_mode("CONSTANT")
        self.cabinet.set_temp(temperature)
        self.switch_activated_buttons("CONSTANT")
        self.input_entry.delete(0, tk.END)
        self.lock.release()

    def power_on(self):
        self.lock.acquire()
        result = self.cabinet.set_mode('STANDBY')
        if result == 'OK':
            messagebox.showinfo("Power On", "Virta on kytketty päälle.")
            self.switch_activated_buttons("STANDBY")
        else:
            messagebox.showerror("Virhe", "Virhe virtakytkennän päälle laittamisessa.")
        self.lock.release()

    def power_off(self):
        self.lock.acquire()
        result = self.cabinet.power_off()
        if result == 'OK':
            messagebox.showinfo("Power Off", "Virta on kytketty pois päältä.")
            self.switch_activated_buttons("OFF")
        else:
            messagebox.showerror("Virhe", "Virhe virtakytkennän pois päältä laittamisessa.")
        self.lock.release()

    def rewrite_program_steps(self):
        program_number = self.input_entry.get()
        if program_number == "":
            program_number = 1  # Using default number 1, if input left empty
            print(f"Creating program {program_number} by default")
            success = self.cabinet.write_program(program_number)
        else:
            self.input_entry.delete(0, tk.END)
            print(f"Creating program {program_number}")
            success = self.cabinet.write_program(program_number)

        if success:
            print(f"PROGRAM {program_number} creation successful")
        else:
            messagebox.showinfo(f"Program (ramp) {program_number} creation failed", f"Try again and make sure the program number {program_number} is not running currently")

    def run_program(self):
        self.lock.acquire()
        program_number = self.input_entry.get()
        if program_number == "":
            print("Running program 1 by default")
            program_number = "1"
        self.switch_activated_buttons("RUN")
        self.cabinet.run_program(program_number)
        self.input_entry.delete(0, tk.END)
        self.lock.release()

    def stop_program(self):
        self.lock.acquire()
        if self.confirmation_messagebox():
            self.cabinet.set_mode("STANDBY")
            self.switch_activated_buttons("STANDBY")
        self.lock.release()

    def start_update_output_thread(self):
        self.output_thread = threading.Thread(target=self.update_output)
        self.output_thread.daemon = True
        self.output_thread.start()

    def close(self):
        self.is_closed = True
        if messagebox.askokcancel("Quit", "Do you want to quit? The cabinet will maintain its mode (program will stay running, etc.). Temp results will not be logged into the csv file"):
            self.window.destroy()

    def switch_activated_buttons(self, cabinet_mode):
        if "OFF" in cabinet_mode:
            self.set_output_off_button.config(state=tk.DISABLED)
            self.set_output_on_button.config(state=tk.NORMAL)
            self.stop_program_button.place_forget()
            self.run_program_button.place(x=440, y=100)
        elif "STANDBY" in cabinet_mode:
            self.set_output_on_button.config(state=tk.DISABLED)
            self.set_output_off_button.config(state=tk.NORMAL)
            self.stop_program_button.place_forget()
            self.run_program_button.place(x=440, y=100)
        elif "RUN" in cabinet_mode:
            self.set_output_on_button.config(state=tk.NORMAL)
            self.set_output_off_button.config(state=tk.NORMAL)
            self.run_program_button.place_forget()
            self.stop_program_button.place(x=440, y=100)
        elif "CONSTANT" in cabinet_mode:
            self.set_output_on_button.config(state=tk.NORMAL)
            self.set_output_off_button.config(state=tk.NORMAL)
            self.stop_program_button.place_forget()
            self.run_program_button.place(x=440, y=100)


    def update_output(self):
        try:
            with open('temp_chamber_results.csv', 'w', newline='', encoding='utf-8') as csvfile:
                writer = csv.writer(csvfile)
                writer.writerow(
                    ['Current', 'Set', 'High', 'Low', 'Mode', 'Time', 'Program number', 'Program step number',
                     "Step duration", "Program cycle count"])
                csvfile.flush()
                counter = 0
                while not self.is_closed:
                    time.sleep(0.3)
                    self.lock.acquire()
                    current_time = datetime.datetime.fromtimestamp(time.time()).strftime("%H:%M:%S")
                    temp_data = self.cabinet.read_temp()
                    cabinet_mode = self.cabinet.query("MODE?")
                    cabinet_mode = cabinet_mode.replace("\r\n", "")
                    if "RUN" in cabinet_mode:
                        program_info = self.cabinet.get_program_status()
                        self.lock.release()
                        self.write_output(temp_data, cabinet_mode, current_time, program_info)
                        if counter % 3 == 0:
                            writer.writerow(
                                [temp_data["current"], program_info["TARGET_TEMP"], temp_data["high"], temp_data["low"],
                                 cabinet_mode, current_time, program_info["PROGRAM_NUMBER"],
                                 program_info["CURRENT_STEP"],
                                 program_info["STEP_TIME_REMAINING"],
                                 re.sub(r'\W+', '', str(program_info["REPEAT_CYCLE_COUNT"]))])
                    else:
                        self.lock.release()
                        self.write_output(temp_data, cabinet_mode, current_time)
                        if counter % 5 == 0:
                            writer.writerow(
                                [temp_data["current"], temp_data["set"], temp_data["high"], temp_data["low"],
                                 cabinet_mode,
                                 current_time])

                    csvfile.flush()
                    counter += 1
        except PermissionError as e:
            messagebox.showinfo(e, "Close the csv file before running the program")
            print(f"Close the csv file before running the program, error: {e}")
            self.window.destroy()

    def write_output(self, temp_data, cabinet_mode, current_time, program_info=None):
        self.output_text.delete('1.0', tk.END)
        self.insert_output_text(f'Current temperature: {temp_data["current"]}\n')
        self.insert_output_text(f'Time: {current_time}\n')
        self.insert_output_text(f'Mode: {cabinet_mode}\n')

        if "RUN" in cabinet_mode:
            self.insert_output_text(f'Current program: {program_info["PROGRAM_NUMBER"]}\n')
            self.insert_output_text(f'Step target temperature: {program_info["TARGET_TEMP"]}\n')
            self.insert_output_text(f'Current step: {program_info["CURRENT_STEP"]}\n')
            self.insert_output_text(f'Current step time left: {program_info["STEP_TIME_REMAINING"]}\n')
            self.insert_output_text(f"Program cycles remaining: {program_info['REPEAT_CYCLE_COUNT']}\n")

        elif "CONSTANT" in cabinet_mode:
            self.insert_output_text(f'Set temperature: {temp_data["set"]}\n')
            self.insert_output_text(f'High: {temp_data["high"]}\n')
            self.insert_output_text(f'Low: {temp_data["low"]}\n')

    def insert_output_text(self, text):
        self.output_text.insert(tk.END, text)

    def confirmation_messagebox(self):
        return messagebox.askokcancel("Quitting the program", "Are you sure you want to stop the program?")


def main():
    temp_cabinet = SU241('GPIB0::10::INSTR')
    mainwindow = MainWindow(temp_cabinet)


if __name__ == "__main__":
    main()
