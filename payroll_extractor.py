import customtkinter
import os
from PIL import Image
import csv
from datetime import datetime, timedelta
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import tkinter as tk
from tkinter import messagebox, scrolledtext, ttk
from selenium.webdriver.chrome.options import Options
import tkinter.font as tkFont
import threading
import requests
import sys
import subprocess
import json
import time
import math
class App(customtkinter.CTk):
    def __init__(self):
        super().__init__()
        self.title("Payroll Data Extractor")
        self.geometry("1100x820+0+0")

        self.grid_rowconfigure(0, weight=1)
        self.grid_columnconfigure(1, weight=1)
        self._is_logged_in = False
        self.is_opened_driver = False
        self.version = "1.1.1"

        image_path = os.path.join(os.path.dirname(os.path.realpath(__file__)), "images")
        self.logo_image = customtkinter.CTkImage(Image.open(os.path.join(image_path, "logo.ico")), size=(40, 40))
        self.iconbitmap(os.path.join(image_path, "logo.ico"))
        self.home_image = customtkinter.CTkImage(light_image=Image.open(os.path.join(image_path, "home_dark.png")), dark_image=Image.open(os.path.join(image_path, "home_light.png")), size=(20, 20))
        self.users_image = customtkinter.CTkImage(light_image=Image.open(os.path.join(image_path, "users_dark.png")), dark_image=Image.open(os.path.join(image_path, "users_light.png")), size=(20, 20))

        self.navigation_frame = customtkinter.CTkFrame(self, corner_radius=0)
        self.navigation_frame.grid(row=0, column=0, sticky="nsew")
        self.navigation_frame.grid_rowconfigure(4, weight=1)

        self.navigation_label = customtkinter.CTkLabel(self.navigation_frame, text="", image=self.logo_image, compound="left", font=customtkinter.CTkFont(size=15, weight="bold"))
        self.navigation_label.grid(row=0, column=0, padx=20, pady=20)

        self.home_button = customtkinter.CTkButton(self.navigation_frame, corner_radius=0, height=40, border_spacing=10, text="Home", fg_color="transparent", text_color=("gray10", "gray90"), hover_color=("gray70", "gray30"), image=self.home_image, anchor="w", command=self.show_home_frame)
        self.home_button.grid(row=1, column=0, sticky="ew")

        self.users_button = customtkinter.CTkButton(self.navigation_frame, corner_radius=0, height=40, border_spacing=10, text="User Data", fg_color="transparent", text_color=("gray10", "gray90"), hover_color=("gray70", "gray30"), image=self.users_image, anchor="w", command=self.show_users_frame)
        self.users_button.grid(row=2, column=0, sticky="ew")

        self.appearance_mode_label = customtkinter.CTkLabel(self.navigation_frame, text="Appearance Mode:", anchor="w")
        self.appearance_mode_label.grid(row=5, column=0)

        self.appearance_mode = customtkinter.CTkOptionMenu(self.navigation_frame, values=["Light", "Dark", "System"], command=self.change_appearance_mode_event)
        self.appearance_mode.grid(row=6, column=0, padx=20, pady=(0, 20), sticky="s")
        customtkinter.set_appearance_mode("Light")
        
        self.read_version()

        self.version_label = customtkinter.CTkLabel(self.navigation_frame, text=f"Version: {self.version}", anchor="w")
        self.version_label.grid(row=7, column=0)

        self.home_frame = customtkinter.CTkFrame(self, corner_radius=0, fg_color="transparent")
        self.home_frame.grid(row=0, column=1, sticky="nsew")

        self.initialize_home_ui()
        self.check_startup()

        self.users_frame = customtkinter.CTkFrame(self, corner_radius=0, fg_color="transparent")
        self.initialize_users_ui()

        self.current_frame = self.home_frame
        self.show_home_frame()

        # self.driver = self.setup_webdriver()
        # self.access_login_page()
        self.protocol("WM_DELETE_WINDOW", self.on_closing)
        self._is_logged_in = False

    def setup_webdriver(self):
        options = Options()
        options.add_argument("--headless")
        options.add_argument("--blink-settings=imagesEnabled=false")
        self.is_opened_driver = True
        return webdriver.Chrome(options=options)

    def access_login_page(self):
        try:
            self.driver.get("")
        except Exception as e:
            self.append_warning(f"Payroll site can’t be reached.")
    def _login(self):
        if self._is_logged_in:
            return True
        if not self.read_credentials():
            self.append_warning("Credentials not found. Please enter and save them.")
            return False
        try:
            wait = WebDriverWait(self.driver, 10)
            username_field = wait.until(EC.visibility_of_element_located((By.ID, "txtUserName")))
            password_field = wait.until(EC.visibility_of_element_located((By.ID, "txtP")))
            login_button = wait.until(EC.visibility_of_element_located((By.ID, "btnLogin")))

            username_field.send_keys(self.user_name)
            password_field.send_keys(self.pwd_payroll)
            login_button.click()
            if self._is_login_successful():
                self._is_logged_in = True
                self.after(0, lambda: self.ui_components.txt_output.insert("end", "Login successful.\n"))
                return True
            else:
                self.after(0, lambda: self.ui_components.txt_output.insert("end", "Login failed. Please check your credentials.\n"))
                return False
        except Exception as e:
            self.append_warning(f"Error during login.")
            return False

    def _is_login_successful(self):
        try:
            # Check for an element that is only present after successful login
            WebDriverWait(self.driver, 20).until(
                EC.presence_of_element_located((By.ID, "header_notification_bar"))
            )
            return True
        except:
            return False
        
    def on_closing(self):
        if hasattr(self, 'driver') and self.driver:
            self.driver.quit()
        self.is_opened_driver = False
        self.destroy()

    def change_appearance_mode_event(self, new_appearance_mode):
        customtkinter.set_appearance_mode(new_appearance_mode)
        self.update_treeview_style()

    def show_home_frame(self):
        self.current_frame.grid_forget()
        self.home_frame.grid(row=0, column=1, sticky="nsew")
        self.current_frame = self.home_frame

    def show_users_frame(self):
        self.current_frame.grid_forget()
        self.users_frame.grid(row=0, column=1, sticky="nsew")
        self.current_frame = self.users_frame
        self.display_csv()
        self.update_users_frame_color()

    def initialize_home_ui(self):
        self.user_name = ""
        self.pwd_payroll = ""
        self.credentials_file = "credentials.txt"
        self.csv_file = "attendance_data.csv"
        self.startup_file_name = "openPayrollExtractor.bat"
        self.startup_folder = os.path.join(os.path.expanduser("~"), "AppData", "Roaming", "Microsoft", "Windows", "Start Menu", "Programs", "Startup")

        self.home_frame.columnconfigure(0, weight=1)
        self.home_frame.columnconfigure(1, weight=1)
        self.home_frame.columnconfigure(2, weight=1)
        self.home_frame.rowconfigure((0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10), weight=0)

        self.text_note = customtkinter.CTkLabel(self.home_frame, text="Welcome to Payroll Extractor. Have a nice working day!", wraplength=400)
        self.text_note.grid(row=0, column=0, columnspan=2, padx=10, pady=5)

        self.txt_username = customtkinter.CTkEntry(self.home_frame, width=300)
        self.txt_username.grid(row=1, column=0, columnspan=2, padx=10, pady=5)
        self.txt_username.delete(0, tk.END)
        self.txt_username.insert(0, self.user_name)

        self.txt_password = customtkinter.CTkEntry(self.home_frame, width=300, show="*")
        self.txt_password.grid(row=2, column=0, columnspan=2, padx=10, pady=5)
        self.txt_password.delete(0, tk.END)
        self.txt_password.insert(0, self.pwd_payroll)
        if self.read_credentials():
            self.txt_username.delete(0, tk.END)
            self.txt_username.insert(0, self.user_name)
            self.txt_password.delete(0, tk.END)
            self.txt_password.insert(0, self.pwd_payroll)

        self.chk_show_password_var = tk.BooleanVar()
        self.chk_show_password = customtkinter.CTkCheckBox(self.home_frame, text="Show Password", variable=self.chk_show_password_var, command=self.toggle_password)
        self.chk_show_password.grid(row=3, column=0, columnspan=2, padx=10, pady=5)

        self.txt_password_visible = customtkinter.CTkEntry(self.home_frame, width=300)
        self.txt_password_visible.grid(row=2, column=0, columnspan=2, padx=10, pady=5)
        self.txt_password_visible.grid_remove()

        self.btn_extract = customtkinter.CTkButton(self.home_frame, text="Extract Data", command=self.start_extract_thread)
        self.btn_extract.grid(row=4, column=0, columnspan=2, padx=10, pady=5)

        self.status_label = customtkinter.CTkLabel(self.home_frame, text="Click the button to extract data")
        self.status_label.grid(row=5, column=0, columnspan=2, padx=10, pady=5)

        output_font = tkFont.Font(family="Courier New", size=13)
        self.txt_output = scrolledtext.ScrolledText(self.home_frame, width=85, height=20, font=output_font)
        self.txt_output.grid(row=6, column=0, columnspan=2, padx=10, pady=5)
        self.txt_output.tag_configure("warning", foreground="red")
        self.read_timedata()

        self.txt_notif = tk.Text(self.home_frame, width=85, height=2, font=output_font)
        self.txt_notif.grid(row=7, column=0, columnspan=2, padx=10, pady=5)
        self.txt_notif.insert(tk.END, "@kien.duong - Release Notes:\nhttp://collab.lge.com/main/display/~kien.duong/Payroll+Data+Extractor+-+Release+Notes")


        self.start_with_windows_var = tk.BooleanVar()
        self.start_with_windows_check = customtkinter.CTkCheckBox(self.home_frame, text="Start with Windows", variable=self.start_with_windows_var, command=self.toggle_startup)
        self.start_with_windows_check.grid(row=8, column=0, columnspan=2, padx=10, pady=5)
        self.btn_update = customtkinter.CTkButton(self.home_frame, text="Check for Update", command=self.check_for_update)
        self.btn_update.grid(row=9, column=0,columnspan=2, padx=10, pady=5)
        self.progress_bar = ttk.Progressbar(self.home_frame, orient=tk.HORIZONTAL, length=300, mode='determinate')
        self.progress_label = tk.Label(self.home_frame, text="")
        self.progress_bar.pack_forget()
        self.progress_label.pack_forget()

    def add_link_to_text(self, text_widget, label, url):
        # Create a tag for the link
        text_widget.tag_configure("link", foreground="blue", underline=True)
        text_widget.tag_bind("link", "<Button-1>", lambda e: self.open_url(url))

        # Insert the label with the tag
        text_widget.insert(tk.END, label, "link")
        
    def initialize_users_ui(self):
        self.tree = ttk.Treeview(self.users_frame, show="headings")
        self.tree.pack(side="top", fill="both", padx=10, pady=10, expand=True)
        self.update_treeview_style()
        self.update_users_frame_color()
        
    def check_for_update(self):
        repo = "duongkien27/Payroll-Extractor"
        api_url = 

        # GitHub token setup
        github_token = ""

        headers = {
            "User-Agent": "Mozilla/5.0",
        }
        if github_token:
            headers["Authorization"] = f"token {github_token}"

        # Check remaining rate limit (optional but good practice)
        rate_limit_url = ""
        try:
            rate_limit_response = requests.get(rate_limit_url, headers=headers)
            rate_limit_response.raise_for_status()
            rate_limit_data = rate_limit_response.json()
            remaining_requests = rate_limit_data['resources']['core']['remaining']

            if remaining_requests == 0:
                messagebox.showwarning("Rate Limit Exceeded", "GitHub API rate limit exceeded. Please try again later.")
                return

            # Proceed with checking for update if rate limit is not exceeded
            response = requests.get(api_url, headers=headers, timeout=10)
            response.raise_for_status()
            data = response.json()

            latest_version = data["tag_name"]

            if latest_version == self.version:
                messagebox.showinfo("Update", "You are using the latest version.")
                return

            answer = messagebox.askyesno("Update Available", f"A new version ({latest_version}) is available.\nDo you want to download and update?")
            if answer:
                self.start_download_thread(data)
        except requests.exceptions.HTTPError as http_err:
            messagebox.showerror("Update Error", f"HTTP error: {http_err}")
        except Exception as e:
            messagebox.showerror("Update Error", f"Failed to check for updates:\n{e}")

    def start_download_thread(self, data):
        self.txt_output.delete(1.0, "end")
        self.txt_output.insert("end", "Downloading...\n")
        self.progress_bar = ttk.Progressbar(self.home_frame, orient=tk.HORIZONTAL, length=300, mode='determinate')
        self.progress_bar.grid(row=10, column=0,columnspan=2, padx=10, pady=5)
        self.progress_label = tk.Label(self.home_frame, text="Downloading: 0%")
        self.progress_label.grid(row=11, column=0,columnspan=2, padx=10, pady=5)
        thread = threading.Thread(target=self.download_update_with_progress, args=(data,))
        thread.daemon = True
        thread.start()

    def download_update_with_progress(self, data):
        # GitHub token setup
        github_token = ""

        headers = {
            "User-Agent": "Mozilla/5.0",
        }
        if github_token:
            headers["Authorization"] = f"token {github_token}"
        latest_version = data["tag_name"]

        asset = next((a for a in data["assets"] if a["name"].endswith(".exe")), None)
        if asset:
            new_exe_name = f"PayrollExtractor_{latest_version}.exe"
            download_url = asset["browser_download_url"]
            download_path = os.path.join(os.getcwd(), new_exe_name)
            total_size = asset["size"]

            self.txt_output.insert(tk.END, f"Downloading update from {download_url}\n")
            self.progress_bar['maximum'] = total_size
            self.progress_bar['value'] = 0

            try:
                with requests.get(download_url, headers=headers, stream=True) as r:
                    r.raise_for_status()
                    downloaded_size = 0
                    with open(download_path, 'wb') as f:
                        for chunk in r.iter_content(chunk_size=8192):
                            f.write(chunk)
                            downloaded_size += len(chunk)
                            percent_complete = int((downloaded_size / total_size) * 100)
                            self.home_frame.after(0, self.update_progress_bar, downloaded_size, f"Downloading: {percent_complete}%")
                self.txt_output.insert(tk.END, f"Downloaded to {download_path}\n")
                self.txt_output.insert(tk.END, f"Launching update...\n")
                self.update_version(latest_version)
                self.remove_from_startup()
                self.launch_new_program(download_path)
                time.sleep(2)
                self.on_closing()
                os._exit(0)
            except requests.exceptions.HTTPError as http_err:
                messagebox.showerror("Download Error", f"HTTP error during download: {http_err}")
            except Exception as e:
                messagebox.showerror("Download Error", f"Failed to download update:\n{e}")
            finally:
                self.home_frame.after(0, self.reset_progress_bar)
        else:
            messagebox.showwarning("Update", "No .exe file found in the latest release.")
            self.home_frame.after(0, self.reset_progress_bar)

    def launch_new_program(self, download_path):
        try:
            subprocess.Popen([download_path])
        except FileNotFoundError:
            messagebox.showerror("Error", f"Executable not found at: {download_path}")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to launch new program: {e}")

    def update_progress_bar(self, value, text):
        self.progress_bar['value'] = value
        self.progress_label['text'] = text

    def reset_progress_bar(self):
        self.progress_bar['value'] = 0
        self.progress_label['text'] = "Download complete" # Or a more appropriate message
   
    def update_treeview_style(self):
        style = ttk.Style()

        style.theme_use("clam")

        if customtkinter.get_appearance_mode() == "Dark":
            style.configure(
                "Treeview",
                background="#333333",
                foreground="white",
                fieldbackground="#333333",
                bordercolor="#333333",
                font=("Arial", 13),
                rowheight=30
            )
            style.configure(
                "Treeview.Heading",
                background="#444444",
                foreground="white",
                font=("Arial", 13, "bold")
            )
        else:
            style.configure(
                "Treeview",
                background="white",
                foreground="black",
                fieldbackground="white",
                bordercolor="white",
                font=("Arial", 13),
                rowheight=30
            )
            style.configure(
                "Treeview.Heading",
                background="lightgray",
                foreground="black",
                font=("Arial", 13, "bold")
            )

    def update_users_frame_color(self):
        if customtkinter.get_appearance_mode() == "Dark":
            self.users_frame.configure(fg_color="#242424")
        else:
            self.users_frame.configure(fg_color="transparent")

    def display_csv(self):
        file_path = "attendance_data.csv"
        if not os.path.exists(file_path):
            messagebox.showwarning("File Not Found", "The file 'attendance_data.csv' was not found!")
            return
        try:
            with open(file_path, newline='', encoding="utf-8") as file:
                reader = csv.reader(file)
                data = list(reader)

                self.tree.delete(*self.tree.get_children())
                self.tree["columns"] = [str(i) for i in range(len(data[0]))]

                for i, column in enumerate(data[0]):
                    self.tree.heading(str(i), text=column)
                    self.tree.column(str(i), width=150, anchor="center")
                self.tree.tag_configure("red_row", foreground="red")
                for row in data[1:]:
                    day = row[1].strip().lower() if len(row) > 1 else ""
                    if day in ["saturday", "sunday","thứ bảy","chủ nhật"]:
                        self.tree.insert("", tk.END, values=row, tags=("red_row",))
                    else:
                        self.tree.insert("", tk.END, values=row)

        except FileNotFoundError:
            self.status_label.configure(text="CSV file not found")

    def toggle_password(self):
        if self.chk_show_password_var.get():
            self.txt_password_visible.delete(0, tk.END)
            self.txt_password_visible.insert(0, self.txt_password.get())
            self.txt_password_visible.grid()
            self.txt_password.grid_remove()
        else:
            self.txt_password.delete(0, tk.END)
            self.txt_password.insert(0, self.txt_password_visible.get())
            self.txt_password_visible.grid_remove()
            self.txt_password.grid()

    def read_credentials(self):
        file_path = "credentials.txt"
        if os.path.exists(file_path):
            with open(file_path, "r") as f:
                lines = f.readlines()
                if len(lines) >= 2:
                    self.user_name = lines[0].strip()
                    self.pwd_payroll = lines[1].strip()
                    return True
        else:
            return False

    def save_credentials(self):
        username = self.txt_username.get()
        if self.chk_show_password_var.get():
            password = self.txt_password_visible.get()
        else:
            password = self.txt_password.get()
        if username and password:
            with open("credentials.txt", "w", encoding="utf-8") as f:
                f.write(f"{username}\n{password}")
            self.user_name = username
            self.pwd_payroll = password
        else:
            messagebox.showwarning("Error", "Please enter valid credentials!")
            
    def save_timedata(self, date, start, end, totalOT, weekendOT, weekdayOT, OTratio):
        data = {
            "date": date,
            "version": self.version,
            "start": start,
            "end": end,
            "totalOT": totalOT,
            "weekendOT": weekendOT,
            "weekdayOT": weekdayOT,
            "OTratio": OTratio
        }

        with open("data.json", "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=4)

    def read_timedata(self):
        file_path = "data.json"

        try:
            with open(file_path, "r", encoding="utf-8") as f:
                data = json.load(f)
                date = data.get("date")
                start = data.get("start")
                end = data.get("end")
                totalOT = data.get("totalOT")
                weekendOT = data.get("weekendOT")
                weekdayOT = data.get("weekdayOT")
                OTratio = data.get("OTratio")
                today = datetime.now().strftime("%d/%m/%Y")
                if not date == today:
                    start = "--:--"
                    end = "--:--"
                self.after(0, lambda: self.txt_output.insert("end", f"--------------------------------------------------\n"))
                self.after(0, lambda: self.txt_output.insert("end", f"Today's Check-in Time: {start}\n"))
                self.after(0, lambda: self.txt_output.insert("end", f"Estimated Check-out Time: {end}\n"))
                self.after(0, lambda: self.txt_output.insert("end", f"--------------------------------------------------\n"))
                self.after(0, lambda: self.txt_output.insert("end", f"Total Overtime This Month: {totalOT} hours\n"))
                self.after(0, lambda: self.txt_output.insert("end", f"Total Weekend Overtime: {weekendOT} hours\n"))
                self.after(0, lambda: self.txt_output.insert("end", f"Total Weekday Overtime: {weekdayOT} hours\n"))
                self.after(0, lambda: self.txt_output.insert("end", f"--------------------------------------------------\n"))
                self.after(0, lambda: self.txt_output.insert("end", f"Current ratio: {OTratio:.2f} (Limit: 2.3)\n"))

        except (FileNotFoundError, json.JSONDecodeError) as e:
            self.status_label.configure(text=f"Error reading time data: {e}")
            return None, None, None, None, None
        

    def read_version(self):
        file_path = "data.json"

        try:
            with open(file_path, "r", encoding="utf-8") as f:
                data = json.load(f)
                self.version = data.get("version")
        except (FileNotFoundError, json.JSONDecodeError):
            self.version = "1.1.1"
        
    def update_version(self, new_version):
        file_path = "data.json"

        # Load existing data
        try:
            with open(file_path, "r", encoding="utf-8") as f:
                data = json.load(f)
        except FileNotFoundError:
            data = {}  # or handle as needed

        # Update the field
        data["version"] = new_version

        # Save the updated data
        with open(file_path, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=4)


    def start_extract_thread(self):
        if not self.is_opened_driver:
            self.driver = self.setup_webdriver()
        self.access_login_page()
        self.txt_output.delete(1.0, "end")
        self.txt_output.insert("end", "Starting Payroll Extraction...\n")
        self.txt_output.insert("end", "Execution time depends on website response. Please be patient!\n")
        self.status_label.configure(text="Starting...")
        self.save_credentials()
        if not self.read_credentials():
            self.status_label.configure(text="Error: Credentials not found!")
            self.txt_output.insert("end", "Please enter and save your credentials first.\n")
            return
        thread = threading.Thread(target=self.extract_data)
        thread.daemon = True
        thread.start()

    def extract_data(self):
        try:
            self.extract_and_save_data(self.user_name, self.pwd_payroll)
            self.after(0, lambda: self.status_label.configure(text="Extraction completed successfully!"))
        except Exception as ex:
            self.after(0, lambda: self.txt_output.insert("end", f"Error: {ex}\n"))
            self.after(0, lambda: self.status_label.configure(text="Error occurred!"))

    def calculate_overtime(self, start_time, end_time, is_weekday):
        try:
            start = datetime.strptime(start_time, "%H:%M")
            end = datetime.strptime(end_time, "%H:%M")
            overtime = end - start

            if is_weekday:
                overtime -= timedelta(hours=9)
            else:
                if end > datetime.strptime("13:15", "%H:%M"):
                    overtime -= timedelta(hours=1)

            overtime_hours = max(0, overtime.total_seconds()) / 3600.0
            return math.floor(overtime_hours * 2) / 2.0

        except ValueError:
            return 0.0

    def append_warning(self, message):
        self.after(0, lambda: self.txt_output.insert("end", f"{message}\n", "warning"))

    def extract_and_save_data(self, username, password):
        try:
            wait = WebDriverWait(self.driver, 30)
            if not self._is_logged_in:
                self._login()
            self.driver.get("example.com")
            wait.until(EC.presence_of_element_located((By.ID, "_ctl0_dtgListctl0")))
            wait.until(EC.presence_of_element_located((By.ID, "header_notification_bar")))

            li_element = wait.until(EC.presence_of_element_located((By.XPATH, "//li[@class='dropdown dropdown-notification dropdown-dark']")))
            p_element = li_element.find_element(By.XPATH, "//p[@style='font-weight: bold']")
            empName = p_element.text
            self.after(0, lambda: self.txt_output.delete(1.0, "end"))
            self.after(0, lambda: self.txt_output.insert("end", f"Hi {empName}! How's it going?\n"))

            table = self.driver.find_element(By.ID, "_ctl0_dtgListctl0")
            rows = table.find_elements(By.TAG_NAME, "tr")

            csv_data = [["Date", "Weekday", "Status", "Start Time", "End Time", "OT"]]

            total_overtime = 0.0
            weekend_overtime = 0.0
            weekday_overtime = 0.0
            today = datetime.now().strftime("%d/%m/%Y")

            for row in rows:
                cells = row.find_elements(By.TAG_NAME, "td")
                if len(cells) >= 5 and cells[1].text.strip():
                    date = cells[1].text.strip()
                    weekday = cells[2].text.strip()
                    status = cells[4].text.strip()

                    gio_bat_dau_input = row.find_element(By.XPATH, ".//input[contains(@id, 'txtGioBatDauTT')]")
                    phut_bat_dau_input = row.find_element(By.XPATH, ".//input[contains(@id, 'txtPhutBatDauTT')]")
                    gio_ket_thuc_input = row.find_element(By.XPATH, ".//input[contains(@id, 'txtGioKetThucTT')]")
                    phut_ket_thuc_input = row.find_element(By.XPATH, ".//input[contains(@id, 'txtPhutKetThucTT')]")

                    start_time = (gio_bat_dau_input.get_attribute("value") or "") + ":" + (phut_bat_dau_input.get_attribute("value") or "")
                    end_time = (gio_ket_thuc_input.get_attribute("value") or "") + ":" + (phut_ket_thuc_input.get_attribute("value") or "")

                    if (start_time != ":"):
                        start_time_cpr = datetime.strptime(start_time, "%H:%M")
                        seven_am = datetime.strptime("07:00", "%H:%M")

                        if start_time_cpr < seven_am:
                            start_time = "07:00"

                    overtime = 0.0
                    if status in ["Danh sách", "List"]:
                        status = "OT"
                        if weekday in ["Thứ Bảy", "Chủ Nhật", "Saturday", "Sunday"]:
                            overtime = self.calculate_overtime(start_time, end_time, 0)
                            weekend_overtime += overtime
                        else:
                            overtime = self.calculate_overtime(start_time, end_time, 1)
                            weekday_overtime += overtime

                        total_overtime += overtime
                    else:
                        status = "Normal"

                    csv_data.append([date, weekday, status, start_time, end_time, overtime])

                    if date == today:
                        self.after(0, lambda: self.txt_output.insert("end", f"--------------------------------------------------\n"))
                        self.after(0, lambda: self.txt_output.insert("end", f"Today's Check-in Time: {start_time}\n"))
                        estimated_end_time = ":"
                        if start_time != ":":
                            estimated = datetime.strptime(start_time, "%H:%M") + timedelta(hours=9)
                            estimated_end_time = estimated.strftime('%H:%M')
                        self.after(0, lambda: self.txt_output.insert("end", f"Estimated Check-out Time: {estimated_end_time}\n"))
                        self.after(0, lambda: self.txt_output.insert("end", f"--------------------------------------------------\n"))
                        self.after(0, lambda: self.txt_output.insert("end", f"Total Overtime This Month: {total_overtime} hours\n"))
                        self.after(0, lambda: self.txt_output.insert("end", f"Total Weekend Overtime: {weekend_overtime} hours\n"))
                        self.after(0, lambda: self.txt_output.insert("end", f"Total Weekday Overtime: {weekday_overtime} hours\n"))
                        self.after(0, lambda: self.txt_output.insert("end", f"--------------------------------------------------\n"))
                        if total_overtime > 40:
                            self.append_warning("⚠ Warning: Your total overtime exceeds 40 hours!\n")

                        ratio_limit = 2.3
                        current_ratio = float('inf')

                        if weekday_overtime > 0:
                            current_ratio = weekend_overtime / weekday_overtime
                        self.after(0, lambda: self.txt_output.insert("end", f"Current ratio: {current_ratio:.2f} (Allowed < {ratio_limit:.2f})"))
                        self.save_timedata(today, start_time, estimated_end_time, total_overtime, weekend_overtime, weekday_overtime, current_ratio)
                        if current_ratio >= ratio_limit:
                            min_weekday_overtime = weekend_overtime / ratio_limit
                            missing_weekday_overtime = min_weekday_overtime - weekday_overtime

                            self.after(0, lambda: self.txt_output.insert("end", f"--------------------------------------------------\n"))
                            self.append_warning("⚠ Warning: You have not met the required overtime ratio!")
                            self.after(0, lambda: self.txt_output.insert("end", f"Minimum weekday overtime required: {min_weekday_overtime:.1f} hours"))
                            self.after(0, lambda: self.txt_output.insert("end", f"You need to work at least {missing_weekday_overtime:.1f} more hours on weekdays."))
                        break

            with open(self.csv_file, "w", newline="", encoding="utf-8") as file:
                writer = csv.writer(file)
                writer.writerows(csv_data)

        except Exception as ex:
            self.append_warning(f"Error: {ex}")

    def add_to_startup(self):
        try:
            exe_path = os.path.abspath(f"PayrollExtractor_{self.version}.exe")
            bat_path = os.path.join(self.startup_folder)
            bat_file_path = os.path.join(bat_path, self.startup_file_name)

            if not os.path.exists(bat_file_path):
                with open(bat_file_path, "w") as bat_file:
                    bat_file.write(f'start "" "{exe_path}"')
                self.status_label.configure(text=f"Added {self.startup_file_name} to startup.")
            else:
                self.status_label.configure(text=f"{self.startup_file_name} already exists in startup.")

        except Exception as e:
            self.status_label.configure(text=f"Error adding to startup: {e}")

    def remove_from_startup(self):
        bat_path = os.path.join(self.startup_folder)
        bat_file_path = os.path.join(bat_path, self.startup_file_name)

        if os.path.exists(bat_file_path):
            try:
                os.remove(bat_file_path)
                self.status_label.configure(text=f"Removed {self.startup_file_name} from startup.")
            except OSError as e:
                self.status_label.configure(text=f"Error removing {self.startup_file_name} from startup: {e}")
        else:
            self.status_label.configure(text=f"{self.startup_file_name} not found in startup folder.")

    def toggle_startup(self):
        if self.start_with_windows_var.get():
            self.add_to_startup()
        else:
            self.remove_from_startup()

    def check_startup(self):
        startup_file_path = os.path.join(self.startup_folder, self.startup_file_name)
        self.start_with_windows_var.set(os.path.exists(startup_file_path))

if __name__ == "__main__":
    app = App()
    app.mainloop()
