import pandas as pd
import matplotlib.pyplot as plt
from tkinter import filedialog, messagebox
import ttkbootstrap as tb
from ttkbootstrap.constants import *
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg

class StepDurationApp(tb.Window):
    def __init__(self):
        super().__init__(themename="flatly")
        self.title("Step Duration Analyzer")
        self.geometry("1000x600")
        
        # UI Elements
        self.label = tb.Label(self, text="Upload Excel File to Analyze Step Durations", font=("Segoe UI", 14))
        self.label.pack(pady=10)

        self.upload_btn = tb.Button(self, text="Upload Excel File", bootstyle=PRIMARY, command=self.load_file)
        self.upload_btn.pack(pady=5)

        self.table_frame = tb.Frame(self)
        self.table_frame.pack(fill=BOTH, expand=True, pady=10)

        self.plot_btn = tb.Button(self, text="Show Graph", bootstyle=SUCCESS, command=self.plot_graph, state="disabled")
        self.plot_btn.pack(pady=5)

        self.df_monthly = None  # Store processed data

    def load_file(self):
        file_path = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx")])
        if not file_path:
            return
        
        try:
            data = pd.read_excel(file_path, sheet_name="Sheet1")
            step_cols = [col for col in data.columns if "Step" in col]

            # Convert Step columns to datetime (time only)
            for col in step_cols:
                data[col] = pd.to_datetime(data[col], format="%H:%M:%S", errors="coerce")

            # Compute durations (Step i+1 - Step i)
            for i in range(len(step_cols) - 1):
                col_name = f"Duration {i+1}"
                data[col_name] = (data[step_cols[i+1]] - data[step_cols[i]]).dt.total_seconds() / 60

            # Extract Month
            data["Month"] = pd.to_datetime(data["Date"]).dt.to_period("M")

            # Compute averages per month
            duration_cols = [col for col in data.columns if "Duration" in col]
            monthly_avg = data.groupby("Month")[duration_cols].mean().reset_index()

            self.df_monthly = monthly_avg

            # Display table
            for widget in self.table_frame.winfo_children():
                widget.destroy()

            tree = tb.Treeview(self.table_frame, columns=["Month"] + duration_cols, show="headings")
            tree.pack(fill=BOTH, expand=True)

            for col in ["Month"] + duration_cols:
                tree.heading(col, text=col)
                tree.column(col, width=120)

            for _, row in monthly_avg.iterrows():
                tree.insert("", "end", values=[row["Month"]] + [round(row[col], 2) for col in duration_cols])

            self.plot_btn.config(state="normal")
        
        except Exception as e:
            messagebox.showerror("Error", f"Failed to process file: {e}")

    def plot_graph(self):
        if self.df_monthly is None:
            return
        
        fig, ax = plt.subplots(figsize=(8, 5))
        duration_cols = [col for col in self.df_monthly.columns if "Duration" in col]

        for col in duration_cols:
            ax.plot(self.df_monthly["Month"].astype(str), self.df_monthly[col], marker="o", label=col)

        ax.set_title("Average Step Duration by Month")
        ax.set_xlabel("Month")
        ax.set_ylabel("Duration (minutes)")
        ax.legend()
        ax.grid(True)

        # Embed plot in Tkinter
        plot_window = tb.Toplevel(self)
        plot_window.title("Step Duration Trends")
        canvas = FigureCanvasTkAgg(fig, master=plot_window)
        canvas.draw()
        canvas.get_tk_widget().pack(fill=BOTH, expand=True)


if __name__ == "__main__":
    app = StepDurationApp()
    app.mainloop()
