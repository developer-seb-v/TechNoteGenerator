import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from tkcalendar import DateEntry
from docx import Document

# Main window setup
root = tk.Tk()
root.title("Sleep Study Tech Note Generator")
root.geometry("750x750")

# Variables for form inputs
vars = {
    "date_of_service": tk.StringVar(),
    "patient_name": tk.StringVar(),
    "dob": tk.StringVar(),
    "referring_md": tk.StringVar(),
    "study_type": tk.StringVar(),
    "study_ran": tk.StringVar(),
    "mask_used": tk.StringVar(),
    "spo2_nadir": tk.StringVar(),
    "sleep_issues": tk.StringVar(value="Unknown"),
    "sleep_tech": tk.StringVar(),
    "location": tk.StringVar()
}

# Helper to create labeled row
def labeled_row(label_text, widget, row):
    tk.Label(root, text=label_text).grid(row=row, column=0, sticky="w", padx=10, pady=5)
    widget.grid(row=row, column=1, padx=10, pady=5)

# Generate narrative and summary
def generate_note():
    try:
        summary_lines = [
            f"Date of Service: {vars['date_of_service'].get()}",
            f"Patient Name: {vars['patient_name'].get()}",
            f"Pt DOB: {vars['dob'].get()}",
            f"Referring MD: {vars['referring_md'].get()}",
            f"Study Type: {vars['study_type'].get()}",
            f"Study Ran: {vars['study_ran'].get()}",
            f"Mask Used: {vars['mask_used'].get()}",
            f"SPO2 Nadir: {vars['spo2_nadir'].get()}%",
            f"Sleep Issues: {vars['sleep_issues'].get()}",
            f"Sleep Tech: {vars['sleep_tech'].get()}",
            f"Location: {vars['location'].get()}"
        ]
        summary_text = "\n".join(summary_lines)

        narrative = (
            f"On {vars['date_of_service'].get()}, {vars['patient_name'].get()} (DOB: {vars['dob'].get()}) "
            f"underwent a {vars['study_type'].get()} at {vars['location'].get()}. The study was conducted "
            f"by {vars['sleep_tech'].get()} using a {vars['mask_used'].get()} mask. The referring provider was "
            f"{vars['referring_md'].get()}. The {vars['study_ran'].get()} was successfully completed. "
            f"SPO2 nadir during the study was {vars['spo2_nadir'].get()}%. Reported sleep issues: "
            f"{vars['sleep_issues'].get()}."
        )

        full_text = summary_text + "\n\n" + narrative
        narrative_text.delete("1.0", tk.END)
        narrative_text.insert(tk.END, full_text)

    except Exception as e:
        messagebox.showerror("Error", str(e))

# Export to Word (.docx)
def export_to_docx():
    try:
        content = narrative_text.get("1.0", tk.END).strip()

        if not content:
            messagebox.showwarning("No Content", "Nothing to export.")
            return

        filepath = filedialog.asksaveasfilename(
            defaultextension=".docx",
            filetypes=[("Word Documents", "*.docx")],
            title="Save As"
        )

        if not filepath:
            return

        doc = Document()
        for line in content.split("\n"):
            doc.add_paragraph(line)
        doc.save(filepath)

        messagebox.showinfo("Export Successful", f"File saved to:\n{filepath}")

    except Exception as e:
        messagebox.showerror("Export Error", str(e))

# --- UI Layout ---

labeled_row("Date of Service", DateEntry(root, textvariable=vars["date_of_service"]), 0)
labeled_row("Patient Name", tk.Entry(root, textvariable=vars["patient_name"]), 1)
labeled_row("Patient DOB", DateEntry(root, textvariable=vars["dob"]), 2)
labeled_row("Referring MD", tk.Entry(root, textvariable=vars["referring_md"]), 3)

study_type_menu = ttk.Combobox(root, textvariable=vars["study_type"], values=["PSG", "CPAP", "BiPAP", "HSAT"])
study_ran_menu = ttk.Combobox(root, textvariable=vars["study_ran"], values=["Full Night", "Split Night", "Titration"])
mask_used_menu = ttk.Combobox(root, textvariable=vars["mask_used"], values=["Nasal", "Full Face", "Nasal Pillows"])
sleep_tech_menu = ttk.Combobox(root, textvariable=vars["sleep_tech"], values=["Tech A", "Tech B", "Tech C"])
location_menu = ttk.Combobox(root, textvariable=vars["location"], values=["Main Lab", "Satellite Office", "Home"])

labeled_row("Study Type", study_type_menu, 4)
labeled_row("Study Ran", study_ran_menu, 5)
labeled_row("Mask Used", mask_used_menu, 6)
labeled_row("SPO2 Nadir (%)", tk.Entry(root, textvariable=vars["spo2_nadir"]), 7)

# Sleep Issues Radio Buttons
tk.Label(root, text="Sleep Issues").grid(row=8, column=0, sticky="w", padx=10)
for i, issue in enumerate(["Yes", "No", "Unknown"]):
    tk.Radiobutton(root, text=issue, variable=vars["sleep_issues"], value=issue).grid(row=8, column=1+i, sticky="w")

labeled_row("Sleep Tech", sleep_tech_menu, 9)
labeled_row("Location", location_menu, 10)

# Buttons
tk.Button(root, text="Generate Tech Note", command=generate_note, bg="#4CAF50", fg="white").grid(row=11, column=0, columnspan=3, pady=10)
tk.Button(root, text="Export to .docx", command=export_to_docx, bg="#2196F3", fg="white").grid(row=12, column=0, columnspan=3)

# Narrative Output
tk.Label(root, text="Generated Narrative:").grid(row=13, column=0, sticky="nw", padx=10)
narrative_text = tk.Text(root, height=10, width=90, wrap="word")
narrative_text.grid(row=14, column=0, columnspan=3, padx=10, pady=10)

root.mainloop()
