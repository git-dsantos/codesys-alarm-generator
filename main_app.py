import tkinter as tk
from tkinter import messagebox, filedialog, ttk
import tkinter.font as tkfont
import os

# Always use UTF-16 LE for CODESYS compatibility
DEFAULT_ENCODING = "utf-16-le"

csv_file_path = None

# Track editing state
editing = False
original_id = None

# Full CSV columns (kept for header creation if file is new)
COLUMNS = [
    "ID",
    "ObservationType",
    "Details1",
    "Details2",
    "Details3",
    "Details4",
    "Details5",
    "Details6",
    "Deactivation",
    "Class",
    "Message",
    "MinPendingTime",
    "Latch1",
    "Latch2",
    "HigherPrioAlarm"
]

# ---------------------------------------------------------------
# UTIL: Manual UTF-16 LE CSV Writer (SAFE FOR CODESYS)
# ---------------------------------------------------------------
def ensure_file_and_header(path):
    """Ensure file exists. If not, create and write header using DEFAULT_ENCODING."""
    if os.path.exists(path):
        return
    header_line = ";".join(COLUMNS) + "\r\n"
    # Create file with header
    with open(path, "w", encoding=DEFAULT_ENCODING) as f:
        f.write(header_line)


def append_row_utf16le(path, row_values):
    """Append a sanitized row to file (manual join, avoids csv.writer with UTF-16)."""
    safe_values = []
    for v in row_values:
        if v is None:
            v = ""
        else:
            v = str(v)
            # remove newlines and trim
            v = v.replace("\n", " ").replace("\r", " ")
            # replace double quotes with single quotes (avoid internal double quotes)
            v = v.replace('"', "'")
        safe_values.append(v)

    line = ";".join(safe_values) + "\r\n"

    # Ensure file exists (and header) before appending
    ensure_file_and_header(path)

    with open(path, "a", encoding=DEFAULT_ENCODING) as f:
        f.write(line)

# ---------------------------------------------------------------
# CREATE CSV
# ---------------------------------------------------------------

BOM_UTF16_LE = b'\xff\xfe'  # UTF-16 LE BOM

def create_new_file():
    global csv_file_path
    file_path = filedialog.asksaveasfilename(
        title="Create New CSV File",
        defaultextension=".csv",
        filetypes=[("CSV Files", "*.csv"), ("All Files", "*.*")]
    )
    if not file_path:
        return  # User canceled

    csv_file_path = file_path
    try:
        with open(csv_file_path, "wb") as f:
            # Write BOM
            f.write(BOM_UTF16_LE)

            # Version line
            version_line = "#Version: 1.0.0.0" + ";" * (len(COLUMNS)-1) + "\r\n"
            f.write(version_line.encode("utf-16-le"))

            # Header line
            header_line = ";".join(COLUMNS) + "\r\n"
            f.write(header_line.encode("utf-16-le"))

    except Exception as e:
        messagebox.showerror("Error", f"Failed to create file:\n{e}")
        return

    label_selected_file.config(text=f"Selected: {csv_file_path}")
    refresh_preview_and_autosize()
    messagebox.showinfo("Success", "New CSV file created successfully!")


# ---------------------------------------------------------------
# READ CSV (simple parser)
# ---------------------------------------------------------------
def read_all_rows(path):
    """Return list of rows (each row is list of columns)."""
    rows = []
    try:
        with open(path, "r", encoding=DEFAULT_ENCODING) as f:
            lines = f.read().splitlines()
    except Exception:
        return rows

    for line in lines:
        if not line.strip():
            continue
        cols = line.split(";")
        # sanitize each cell
        cols = [c.strip().strip('"').strip("'") for c in cols]
        # pad to full length
        while len(cols) < len(COLUMNS):
            cols.append("")
        rows.append(cols)
    return rows


# ---------------------------------------------------------------
# PREVIEW / TREE HELPERS
# ---------------------------------------------------------------
def build_condition(details1, details2, details3):
    # Add spaces around Details2 so it looks like "A = TRUE"
    details1 = details1 or ""
    details2 = details2 or ""
    details3 = details3 or ""
    return f"{details1} {details2} {details3}".strip()


def try_int_key(x):
    """Return tuple for sorting: numeric if possible, else lexicographic."""
    try:
        return (0, int(x))
    except Exception:
        return (1, str(x))


def refresh_preview_and_autosize():
    # Clear tree
    for row in tree.get_children():
        tree.delete(row)

    if not csv_file_path:
        return

    rows = read_all_rows(csv_file_path)
    # Skip first two rows (version + header)
    if len(rows) > 2:
        data_rows = rows[2:]
    else:
        data_rows = []

    # Sort by ID (try numeric)
    data_rows_sorted = sorted(data_rows, key=lambda r: try_int_key(r[0] if r and len(r) > 0 else ""))

    # Populate tree with ID, Condition, Class, Message
    for r in data_rows_sorted:
        id_val = r[0]
        details1 = r[2] if len(r) > 2 else ""
        details2 = r[3] if len(r) > 3 else "="
        details3 = r[4] if len(r) > 4 else ""
        class_val = r[9] if len(r) > 9 else ""
        message = r[10] if len(r) > 10 else ""
        condition = build_condition(details1, details2, details3)
        tree.insert("", "end", values=(id_val, condition, class_val, message))

    autosize_columns()


def autosize_columns():
    # Use the app font defined for your app
    font = app_font

    # Iterate over all columns
    for col in columns_display:
        # Start with the header width
        max_width = font.measure(col)

        # Measure all values in this column
        for item in tree.get_children():
            val = str(tree.item(item, "values")[columns_display.index(col)])
            w = font.measure(val)
            if w > max_width:
                max_width = w

        # Set the column width with some padding
        tree.column(col, width=max_width + 20)



# ---------------------------------------------------------------
# ID DUPLICATE CHECK
# ---------------------------------------------------------------
def id_exists(target_id, exclude_id=None):
    """Check if ID exists in CSV, optionally excluding one ID (useful when editing)."""
    if not csv_file_path:
        return False
    rows = read_all_rows(csv_file_path)
    # skip first two rows
    if len(rows) > 2:
        data_rows = rows[2:]
    else:
        data_rows = []
    for r in data_rows:
        if len(r) > 0:
            if r[0] == target_id:
                if exclude_id is not None and r[0] == exclude_id:
                    continue
                return True
    return False


# ---------------------------------------------------------------
# DELETE SELECTED ROW
# ---------------------------------------------------------------
def delete_selected():
    global editing, original_id
    selected = tree.selection()
    if not selected:
        messagebox.showinfo("No selection", "Please select a row to delete.")
        return

    item = selected[0]
    vals = tree.item(item, "values")
    target_id = vals[0]

    if not messagebox.askyesno("Confirm delete", f"Delete entry ID '{target_id}'?"):
        return

    # Read all rows and remove the first matching ID (skip first two rows)
    rows = read_all_rows(csv_file_path)
    new_lines = []
    removed = False

    if len(rows) > 2:
        new_lines.extend(rows[:2])  # keep first two rows
        data_rows = rows[2:]
    else:
        data_rows = rows

    for r in data_rows:
        if not removed and len(r) > 0 and r[0] == target_id:
            removed = True
            continue
        new_lines.append(r)

    if not removed:
        messagebox.showerror("Not found", "Selected ID not found in file.")
        return

    # Write back all lines (overwrite) using UTF-16 LE
    try:
        with open(csv_file_path, "w", encoding=DEFAULT_ENCODING) as f:
            for row in new_lines:
                while len(row) < len(COLUMNS):
                    row.append("")
                line = ";".join(row) + "\r\n"
                f.write(line)
    except Exception as e:
        messagebox.showerror("Write Error", f"Failed to update CSV:\n{e}")
        return

    if editing and original_id == target_id:
        editing = False
        original_id = None
        btn_add.config(state=tk.NORMAL)
        btn_save.config(state=tk.DISABLED)
        entry_id.config(state=tk.NORMAL)

    refresh_preview_and_autosize()
    messagebox.showinfo("Deleted", f"Entry ID '{target_id}' removed.")


# ---------------------------------------------------------------
# FILE BROWSER & ADD / EDIT ENTRY
# ---------------------------------------------------------------
def browse_file():
    global csv_file_path
    file_path = filedialog.askopenfilename(title="Select CSV File", filetypes=[("CSV Files", "*.csv"), ("All Files", "*.*")])
    if file_path:
        csv_file_path = file_path
        label_selected_file.config(text=f"Selected: {file_path}")
        ensure_file_and_header(csv_file_path)
        refresh_preview_and_autosize()
        messagebox.showinfo("File Loaded", "CSV file loaded successfully!")


def add_entry():
    global csv_file_path
    if not csv_file_path:
        messagebox.showerror("Error", "Please select a CSV file first.")
        return

    id_value = entry_id.get().strip()
    details1 = entry_details1.get().strip()
    details2 = "="
    details3 = combo_details3.get()
    observation_type = combo_obs_type.get()
    class_value = combo_class.get()
    message = text_message.get("1.0", tk.END).strip()

    if not id_value:
        messagebox.showerror("Error", "ID is required.")
        return

    if id_exists(id_value):
        messagebox.showerror("Duplicate ID", f"ID '{id_value}' already exists!")
        return

    row = [
        id_value,
        observation_type,
        details1,
        details2,
        details3,
        "",
        "",
        "",
        "",
        class_value,
        message,
        "",
        "",
        "",
        ""
    ]

    try:
        append_row_utf16le(csv_file_path, row)
    except Exception as e:
        messagebox.showerror("Write Error", f"Failed to write to CSV:\n{e}")
        return

    # clear form
    entry_id.delete(0, tk.END)
    entry_details1.delete(0, tk.END)
    combo_details3.set("TRUE")
    combo_class.set("Error")
    combo_obs_type.set("Digital")
    text_message.delete("1.0", tk.END)

    refresh_preview_and_autosize()
    messagebox.showinfo("Success", "Entry added successfully!")


# ---------------------------------------------------------------
# EDIT SELECTED / SAVE CHANGES
# ---------------------------------------------------------------
def edit_selected():
    global editing, original_id
    selected = tree.selection()
    if not selected:
        messagebox.showinfo("No selection", "Please select a row to edit.")
        return

    item = selected[0]
    vals = tree.item(item, "values")
    target_id = vals[0]

    rows = read_all_rows(csv_file_path)
    data_rows = rows[2:] if len(rows) > 2 else rows

    found = None
    for r in data_rows:
        if len(r) > 0 and r[0] == target_id:
            found = r
            break

    if not found:
        messagebox.showerror("Not found", "Selected ID not found in file.")
        return

    # Populate form with values
    entry_id.delete(0, tk.END)
    entry_id.insert(0, found[0])
    entry_details1.delete(0, tk.END)
    entry_details1.insert(0, found[2] if len(found) > 2 else "")
    combo_obs_type.set(found[1] if len(found) > 1 else "Digital")
    combo_details3.set(found[4] if len(found) > 4 else "TRUE")
    combo_class.set(found[9] if len(found) > 9 else "Error")
    text_message.delete("1.0", tk.END)
    text_message.insert("1.0", found[10] if len(found) > 10 else "")

    # Enter edit mode
    editing = True
    original_id = found[0]
    btn_add.config(state=tk.DISABLED)
    btn_save.config(state=tk.NORMAL)
    entry_id.config(state=tk.NORMAL)


def save_changes():
    global editing, original_id
    if not editing:
        return

    new_id = entry_id.get().strip()
    details1 = entry_details1.get().strip()
    details2 = "="
    details3 = combo_details3.get()
    observation_type = combo_obs_type.get()
    class_value = combo_class.get()
    message = text_message.get("1.0", tk.END).strip()

    if not new_id:
        messagebox.showerror("Error", "ID is required.")
        return

    if id_exists(new_id, exclude_id=original_id):
        messagebox.showerror("Duplicate ID", f"ID '{new_id}' already exists!")
        return

    rows = read_all_rows(csv_file_path)
    header_rows = rows[:2] if len(rows) > 2 else []
    data_rows = rows[2:] if len(rows) > 2 else rows

    updated = False
    new_all = header_rows.copy()

    for r in data_rows:
        if not updated and len(r) > 0 and r[0] == original_id:
            new_r = [
                new_id,
                observation_type,
                details1,
                details2,
                details3,
                "",
                "",
                "",
                "",
                class_value,
                message,
                "",
                "",
                "",
                ""
            ]
            new_all.append(new_r)
            updated = True
        else:
            new_all.append(r)

    if not updated:
        messagebox.showerror("Not found", "Original ID not found in file; cannot save changes.")
        return

    try:
        with open(csv_file_path, "w", encoding=DEFAULT_ENCODING) as f:
            for row in new_all:
                while len(row) < len(COLUMNS):
                    row.append("")
                line = ";".join(row) + "\r\n"
                f.write(line)
    except Exception as e:
        messagebox.showerror("Write Error", f"Failed to update CSV:\n{e}")
        return

    editing = False
    original_id = None
    btn_add.config(state=tk.NORMAL)
    btn_save.config(state=tk.DISABLED)

    # clear form
    entry_id.delete(0, tk.END)
    entry_details1.delete(0, tk.END)
    combo_details3.set("TRUE")
    combo_class.set("Error")
    combo_obs_type.set("Digital")
    text_message.delete("1.0", tk.END)

    refresh_preview_and_autosize()
    messagebox.showinfo("Saved", "Changes saved successfully!")


# ---------------------------------------------------------------
# GUI SETUP
# ---------------------------------------------------------------
root = tk.Tk()
root.title("CODESYS Alarm Editor")

# Maximize window on startup
try:
    root.state('zoomed')  # Windows
except:
    root.attributes('-zoomed', True)  # Linux / some systems

root.minsize(900, 500)

# Define a consistent font
app_font = tkfont.Font(family="Arial", size=10)

# Top: file selection
top_frame = tk.Frame(root, padx=10, pady=8)
top_frame.pack(fill=tk.X)

tk.Label(top_frame, text="Select CSV File:", font=("Arial", 10, "bold")).pack(side=tk.LEFT)
btn_browse = tk.Button(top_frame, text="Browse...", width=18, command=browse_file)
btn_browse.pack(side=tk.LEFT, padx=8)
btn_create = tk.Button(top_frame, text="Create New File...", width=18, command=create_new_file)
btn_create.pack(side=tk.LEFT, padx=8)
label_selected_file = tk.Label(top_frame, text="No file selected", fg="gray")
label_selected_file.pack(side=tk.LEFT, padx=8)

# Main split: form (left) + preview (right)
main_frame = tk.Frame(root, padx=10, pady=10)
main_frame.pack(fill=tk.BOTH, expand=True)

form_frame = tk.Frame(main_frame)
form_frame.pack(side=tk.LEFT, fill=tk.Y, padx=(0, 20))

preview_frame = tk.Frame(main_frame)
preview_frame.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True)

# Form fields
lbl = lambda t: tk.Label(form_frame, text=t)
lbl("ID:").pack(anchor="w")
entry_id = tk.Entry(form_frame, width=30)
entry_id.pack(anchor="w", pady=3)

lbl("Details1:").pack(anchor="w")
entry_details1 = tk.Entry(form_frame, width=30)
entry_details1.pack(anchor="w", pady=3)

lbl("ObservationType:").pack(anchor="w")
combo_obs_type = ttk.Combobox(form_frame, values=["Digital", "Analog", "Other"], state="readonly", width=27)
combo_obs_type.set("Digital")
combo_obs_type.pack(anchor="w", pady=3)

lbl("Details3:").pack(anchor="w")
combo_details3 = ttk.Combobox(form_frame, values=["TRUE", "FALSE"], state="readonly", width=27)
combo_details3.set("TRUE")
combo_details3.pack(anchor="w", pady=3)

lbl("Class:").pack(anchor="w")
combo_class = ttk.Combobox(form_frame, values=["Error", "Warning", "Info"], state="readonly", width=27)
combo_class.set("Error")
combo_class.pack(anchor="w", pady=3)

lbl("Message:").pack(anchor="w")
text_message = tk.Text(form_frame, width=40, height=6, font=app_font)
text_message.pack(anchor="w", pady=3)

btn_add = tk.Button(form_frame, text="Add Entry", width=20, height=2, command=add_entry)
btn_add.pack(pady=12)

btn_delete = tk.Button(form_frame, text="Delete Selected", width=20, height=2, command=delete_selected)
btn_delete.pack(pady=4)

btn_edit = tk.Button(form_frame, text="Edit Selected", width=20, height=2, command=edit_selected)
btn_edit.pack(pady=4)

btn_save = tk.Button(form_frame, text="Save Changes", width=20, height=2, command=save_changes)
btn_save.pack(pady=4)
btn_save.config(state=tk.DISABLED)

# Preview Treeview with columns: ID | Condition | Class | Message
columns_display = ("ID", "Condition", "Class", "Message")

tree = ttk.Treeview(preview_frame, columns=columns_display, show="headings")
for col in columns_display:
    tree.heading(col, text=col)
    tree.column(col, width=150, anchor="w")

vsb = ttk.Scrollbar(preview_frame, orient="vertical", command=tree.yview)
tree.configure(yscrollcommand=vsb.set)
vsb.pack(side=tk.RIGHT, fill=tk.Y)

tree.pack(fill=tk.BOTH, expand=True)

refresh_preview_and_autosize()

root.mainloop()
