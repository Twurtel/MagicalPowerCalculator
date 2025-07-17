import math
import re
import pandas as pd
import tkinter as tk
from tkinter import messagebox

# ---------- Helper to clean up hex codes ----------
def sanitize_color(code):
    """
    Ensures `code` is "#RRGGBB" (upper‐case).
    If invalid, returns magenta "#FF00BF" so you can spot errors.
    """
    if not isinstance(code, str):
        return "#FF00BF"
    s = code.strip()
    m = re.match(r'^(#?)([0-9A-Fa-f]{6})$', s)
    if m:
        return "#" + m.group(2).upper()
    return "#FF00BF"

# ---------- Load & prepare data ----------
xls = pd.ExcelFile("Magical Power Spreadsheet.xlsx")

# Stat multipliers + Display Color column
stat_df = pd.read_excel(xls, "Stat multipliers")
stat_df.columns = stat_df.columns.str.strip()

stat_names       = stat_df["Stat:"].str.strip().tolist()
stat_multipliers = dict(zip(
    stat_names,
    stat_df["Base Stat Multiplier:"]
))

raw_colors = stat_df["Display Color"].astype(str).str.strip()
stat_colors = {
    stat: sanitize_color(col)
    for stat, col in zip(stat_names, raw_colors)
}

# Powers sheet
powers_df        = pd.read_excel(xls, "Powers").fillna(0)
powerstone_names = powers_df["Power Name"].tolist()

# Stat calculation function
def calculate_stat(base, mult, mp):
    ln = math.log(1 + 0.0019 * mp)
    return (base / 100) * mult * 719.28 * (ln ** 1.2)

# ---------- GUI ----------
def launch_gui():
    root = tk.Tk()
    root.title("Magical Powerstone Stat Calculator")
    root.geometry("600x600")
    root.configure(bg="black")

    # 1) Generate safe tag IDs (no spaces)
    tag_ids = {
        stat: stat.lower().replace(" ", "_")
        for stat in stat_names
    }

    # 2) Top control bar
    ctrl = tk.Frame(root, bg="black")
    ctrl.pack(fill="x", padx=10, pady=10)

    tk.Label(ctrl, text="Select Powerstone:", fg="white", bg="black")\
      .grid(row=0, column=0, sticky="w")
    power_var = tk.StringVar(value=powerstone_names[0])
    om = tk.OptionMenu(ctrl, power_var, *powerstone_names)
    om.config(bg="black", fg="white",
              highlightthickness=0,
              activebackground="gray", activeforeground="white")
    om["menu"].config(bg="black", fg="white")
    om.grid(row=0, column=1, sticky="w", padx=(5, 20))

    tk.Label(ctrl, text="Magical Power:", fg="white", bg="black")\
      .grid(row=0, column=2, sticky="w")
    mp_entry = tk.Entry(ctrl, bg="black", fg="white",
                        insertbackground="white", width=8)
    mp_entry.insert(0, "1000")
    mp_entry.grid(row=0, column=3, sticky="w", padx=5)

    # 3) Text area for results
    output_text = tk.Text(root,
                          wrap="word", height=25,
                          bg="black", fg="white",
                          insertbackground="white",
                          state="disabled")
    output_text.pack(fill="both", expand=True, padx=10, pady=(0, 10))

    # 4) Configure all tags upfront
    for stat, tag in tag_ids.items():
        color = stat_colors.get(stat, "#FF00BF")
        output_text.tag_config(tag, foreground=color)

    # 5) Calculation handler
    def on_calculate(*args):
        try:
            sel = power_var.get()
            mp  = float(mp_entry.get())
            row = powers_df[powers_df["Power Name"] == sel].iloc[0]

            output_text.config(state="normal")
            output_text.delete("1.0", "end")
            output_text.insert("end",
                f"Results for '{sel}' with MP = {mp:.0f}:\n\n"
            )

            for stat, mult in stat_multipliers.items():
                base = row.get(stat, 0)
                val  = calculate_stat(base, mult, mp)

                # SKIP zero values
                if val == 0:
                    continue

                tag = tag_ids[stat]
                output_text.insert("end",
                    f"  • {stat}: {val:.2f}\n",
                    tag
                )

            ub = row.get("Unique Power Bonus", 0)
            if ub:
                output_text.insert("end",
                    f"\n  ✦ Unique Power Bonus: {ub}\n"
                )

            output_text.config(state="disabled")

        except Exception as e:
            messagebox.showerror("Error", f"Calculation failed:\n{e}")

    # 6) Bind dropdown & entry to auto-calc
    power_var.trace_add("write", on_calculate)
    mp_entry.bind("<Return>", lambda e: on_calculate())

    # 7) Initial display
    on_calculate()

    root.mainloop()

if __name__ == "__main__":
    launch_gui()

