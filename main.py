import tkinter as tk
from tkinter import ttk, filedialog
import csv
from datetime import datetime

# Room line object
class RoomEntry(tk.Frame):
    def __init__(self, parent, update_callback, data=None):
        super().__init__(parent)
        self.configure(bg="lightgray")

        self.update_callback = update_callback
        self.completion_time = data[6] if data and data[6] != "Not Completed" else None

        # Room Name
        self.room_name = tk.StringVar(value=data[0] if data else "")
        tk.Entry(self, textvariable=self.room_name, width=15).grid(row=0, column=0, padx=10, pady=5)

        # Room Type
        self.room_type = tk.StringVar(value=data[1] if data else "")
        tk.Entry(self, textvariable=self.room_type, width=15).grid(row=0, column=1, padx=10, pady=5)

        # Door Style
        self.door_style = tk.StringVar(value=data[2] if data else "")
        tk.Entry(self, textvariable=self.door_style, width=15).grid(row=0, column=2, padx=10, pady=5)

        # Checkboxes
        self.check_var1 = tk.BooleanVar(value=(data[3] == "Yes") if data else False)
        self.check_var2 = tk.BooleanVar(value=(data[4] == "Yes") if data else False)

        self.check1 = tk.Checkbutton(self, variable=self.check_var1, command=self.update_status)
        self.check1.grid(row=0, column=3, padx=10, pady=5)

        self.check2 = tk.Checkbutton(self, variable=self.check_var2, command=self.update_status)
        self.check2.grid(row=0, column=4, padx=10, pady=5)

        # Door Count
        self.door_count = tk.StringVar(value=data[5] if data else "")
        tk.Entry(self, textvariable=self.door_count, width=10).grid(row=0, column=5, padx=10, pady=5)

        self.update_status()

    def update_status(self):
        """Update background color and set timestamp when completed."""
        if self.check_var1.get() and self.check_var2.get():
            self.configure(bg="lightgreen")
            if not self.completion_time:
                self.completion_time = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        else:
            self.configure(bg="lightgray")
            self.completion_time = None  
        self.update_callback()

    def get_data(self):
        """Return room details for saving."""
        return [
            self.room_name.get(),
            self.room_type.get(),
            self.door_style.get(),
            "Yes" if self.check_var1.get() else "No",
            "Yes" if self.check_var2.get() else "No",
            self.door_count.get(),
            self.completion_time if self.completion_time else "Not Completed"
        ]

# Main app window
class WorkOrderApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Work Order Entry")
        self.root.geometry("575x500")

        # Work Order & Project Name
        tk.Label(root, text="Work Order #:").grid(row=0, column=0, padx=5, pady=5)
        self.work_order_entry = tk.Entry(root)
        self.work_order_entry.grid(row=0, column=1, padx=5, pady=5)

        tk.Label(root, text="Project Name:").grid(row=0, column=2, padx=5, pady=5)
        self.project_name_entry = tk.Entry(root)
        self.project_name_entry.grid(row=0, column=3, padx=5, pady=5)

        # Room List Frame
        self.room_frame = tk.Frame(root)
        self.room_frame.grid(row=1, column=0, columnspan=4, padx=10, pady=10)

        # Room Column Titles
        tk.Label(self.room_frame, text="Room Name").grid(row=0, column=0, padx=20, pady=2)
        tk.Label(self.room_frame, text="Room Type").grid(row=0, column=1, padx=25, pady=2)
        tk.Label(self.room_frame, text="Door Style").grid(row=0, column=2, padx=25, pady=2)
        tk.Label(self.room_frame, text="Nests").grid(row=0, column=3, padx=7, pady=2)
        tk.Label(self.room_frame, text="Labels").grid(row=0, column=4, padx=7, pady=2)
        tk.Label(self.room_frame, text="Door Count").grid(row=0, column=5, padx=15, pady=2)

        self.rooms = []

        # Buttons
        self.add_room_button = tk.Button(root, text="Add Room", command=self.add_room)
        self.add_room_button.grid(row=2, column=0, columnspan=4, pady=5)

        self.save_button = tk.Button(root, text="Save to CSV", command=self.save_to_csv)
        self.save_button.grid(row=3, column=0, columnspan=4, pady=5)

        self.load_button = tk.Button(root, text="Load CSV", command=self.load_from_csv)
        self.load_button.grid(row=4, column=0, columnspan=4, pady=5)

    def add_room(self, data=None):
        """Add a new room entry."""
        room_entry = RoomEntry(self.room_frame, self.update_ui, data)
        room_entry.grid(row=len(self.rooms) + 1, column=0, columnspan=6, pady=2, sticky="ew")
        self.rooms.append(room_entry)

    def update_ui(self):
        """Update UI when checkboxes change."""
        pass

    def save_to_csv(self):
        """Save all data to a CSV file."""
        work_order = self.work_order_entry.get()
        project_name = self.project_name_entry.get()

        if not work_order or not project_name:
            print("Please enter Work Order # and Project Name before saving.")
            return

        filename = filedialog.asksaveasfilename(defaultextension=".csv", filetypes=[("CSV files", "*.csv")])
        if not filename:
            return

        with open(filename, mode="w", newline="") as file:
            writer = csv.writer(file)
            writer.writerow(["Work Order #", "Project Name", "Room Name", "Room Type", "Door Style", "Nests", "Labels", "Door Count", "Completion Time"])

            for room in self.rooms:
                writer.writerow([work_order, project_name] + room.get_data())

        print(f"Data saved to {filename} successfully!")

    def load_from_csv(self):
        """Load data from a CSV file and reset the GUI."""
        filename = filedialog.askopenfilename(filetypes=[("CSV files", "*.csv")])
        if not filename:
            return

        with open(filename, mode="r") as file:
            reader = csv.reader(file)
            rows = list(reader)

            if len(rows) < 2:
                print("Invalid CSV format.")
                return

            # Clear existing fields
            self.work_order_entry.delete(0, tk.END)
            self.project_name_entry.delete(0, tk.END)

            # Clear existing rooms
            for room in self.rooms:
                room.destroy()
            self.rooms.clear()

            # Load Work Order & Project Name
            first_row = rows[1]
            self.work_order_entry.insert(0, first_row[0])
            self.project_name_entry.insert(0, first_row[1])

            # Load rooms
            for row in rows[1:]:
                self.add_room(row[2:])

        print(f"Data loaded from {filename} successfully!")

        def clear_blank_rooms(self):
            # Clear all rooms with no room_name
            for room in self.rooms:
                if room.get_data = None:
                    room.destroy()


if __name__ == "__main__":
    root = tk.Tk()
    app = WorkOrderApp(root)
    root.mainloop()
