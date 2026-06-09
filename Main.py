
import customtkinter as ctk
from tkinter import filedialog, messagebox
from tksheet import Sheet
import pandas as pd
import os

from compare import compare_excel, load_excel


class ExcelComparisonTool:

    def __init__(self, root):
        self.root = root
        self.root.title("SheetDiff")
        self.root.geometry("1600x900")

        self.root.after(
            100,
            lambda: self.root.state("zoomed")
        )

        self.file_a = ""
        self.file_b = ""

        self.df_a = None
        self.df_b = None
        self.differences = pd.DataFrame()

        self.build_ui()

    def build_ui(self):

        title = ctk.CTkLabel(
            self.root,
            text="SheetDiff",
            font=("Segoe UI", 28, "bold")
        )
        title.pack(pady=10)

        toolbar = ctk.CTkFrame(self.root)
        toolbar.pack(fill="x", padx=10, pady=5)

        self.file_a_entry = ctk.CTkEntry(toolbar, width=600)
        self.file_a_entry.grid(row=0, column=0, padx=5, pady=5)

        ctk.CTkButton(
            toolbar,
            text="Browse File A",
            command=self.select_file_a
        ).grid(row=0, column=1, padx=5)

        self.file_b_entry = ctk.CTkEntry(toolbar, width=600)
        self.file_b_entry.grid(row=0, column=2, padx=5, pady=5)

        ctk.CTkButton(
            toolbar,
            text="Browse File B",
            command=self.select_file_b
        ).grid(row=0, column=3, padx=5)

        ctk.CTkLabel(
            toolbar,
            text="Key Column:"
        ).grid(row=0, column=4, padx=5)

        self.key_option = ctk.CTkOptionMenu(
            toolbar,
            values=[""]
        )

        self.key_option.grid(
            row=0,
            column=5,
            padx=5
        )

        ctk.CTkButton(
            toolbar,
            text="Compare",
            command=self.compare_files
        ).grid(row=0, column=4, padx=10)

        ctk.CTkButton(
            toolbar,
            text="Export Report",
            command=self.export_report
        ).grid(row=0, column=5, padx=5)

        ctk.CTkButton(
            toolbar,
            text="Clear",
            command=self.clear_all
        ).grid(row=0, column=6, padx=5)

        self.theme_switch = ctk.CTkSwitch(
            toolbar,
            text="Dark Mode",
            command=self.toggle_theme
        )
        self.theme_switch.select()
        self.theme_switch.grid(row=0, column=7, padx=10)

        sheet_frame = ctk.CTkFrame(self.root)
        sheet_frame.pack(fill="both", expand=True, padx=10, pady=10)

        left_frame = ctk.CTkFrame(sheet_frame)
        left_frame.pack(side="left", fill="both", expand=True, padx=(0, 5))

        right_frame = ctk.CTkFrame(sheet_frame)
        right_frame.pack(side="left", fill="both", expand=True, padx=(5, 0))

        self.left_title = ctk.CTkLabel(
            left_frame,
            text="File A",
            font=("Segoe UI", 16, "bold")
        )
        self.left_title.pack(pady=5)

        self.left_sheet = Sheet(left_frame)
        self.left_sheet.pack(fill="both", expand=True, padx=5, pady=5)

        self.right_title = ctk.CTkLabel(
            right_frame,
            text="File B",
            font=("Segoe UI", 16, "bold")
        )
        self.right_title.pack(pady=5)

        self.right_sheet = Sheet(right_frame)
        self.right_sheet.pack(fill="both", expand=True, padx=5, pady=5)

        self.status_label = ctk.CTkLabel(
            self.root,
            text="Ready",
            anchor="w"
        )
        self.status_label.pack(fill="x", padx=10, pady=5)

    def select_file_a(self):

        file_path = filedialog.askopenfilename(
            filetypes=[("Excel Files", "*.xlsx *.xls")]
        )

        if file_path:
            self.file_a = file_path

            self.file_a_entry.delete(0, "end")
            self.file_a_entry.insert(0, file_path)

            self.load_left_sheet()

    def select_file_b(self):

        file_path = filedialog.askopenfilename(
            filetypes=[("Excel Files", "*.xlsx *.xls")]
        )

        if file_path:
            self.file_b = file_path

            self.file_b_entry.delete(0, "end")
            self.file_b_entry.insert(0, file_path)

            self.load_right_sheet()

    def load_left_sheet(self):

        try:
            self.df_a = load_excel(self.file_a)

            filename = os.path.basename(self.file_a)

            self.left_title.configure(
                text=filename
            )

            self.left_sheet.headers(
                list(self.df_a.columns)
            )

            self.left_sheet.set_sheet_data(
                self.df_a.values.tolist()
            )

            columns = list(self.df_a.columns)

            self.key_option.configure(
                values=columns
            )

            if columns:
                self.key_option.set(columns[0])

        except Exception as e:
            messagebox.showerror("Error", str(e))

    def load_right_sheet(self):

        try:
            self.df_b = load_excel(self.file_b)

            filename = os.path.basename(self.file_b)

            self.right_title.configure(
                text=filename
            )

            self.right_sheet.headers(
                list(self.df_b.columns)
            )

            self.right_sheet.set_sheet_data(
                self.df_b.values.tolist()
            )

            columns = list(self.df_b.columns)

            self.key_option.configure(
                values=columns
            )

            if columns:
                self.key_option.set(columns[0])

        except Exception as e:
            messagebox.showerror("Error", str(e))

    def compare_files(self):

        if not self.file_a or not self.file_b:
            messagebox.showerror(
                "Error",
                "Please select both files."
            )
            return

        try:

            self.status_label.configure(
                text="Comparing files..."
            )

            self.root.update()

            key_column = self.key_option.get()

            (
                self.differences,
                self.display_df_a,
                self.display_df_b
            ) = compare_excel(
                self.file_a,
                self.file_b,
                key_column
            )

            self.highlight_differences()

            self.status_label.configure(
                text=f"Differences Found: {len(self.differences)}"
            )

            messagebox.showinfo(
                "Complete",
                f"{len(self.differences)} differences found."
            )

        except Exception as e:
            messagebox.showerror("Error", str(e))

    def highlight_differences(self):

        self.clear_highlights()

        def normalize_key(value):

            if pd.isna(value):
                return ""

            try:
                return str(int(float(value)))
            except Exception:
                return str(value).strip()

        try:

            # Reload sheets
            self.left_sheet.headers(
                list(self.display_df_a.columns)
            )

            self.right_sheet.headers(
                list(self.display_df_b.columns)
            )

            self.left_sheet.set_sheet_data(
                self.display_df_a.fillna("").values.tolist()
            )

            self.right_sheet.set_sheet_data(
                self.display_df_b.fillna("").values.tolist()
            )

            key_column = self.key_option.get()

            # Build lookup tables
            left_key_map = {
                normalize_key(v): idx
                for idx, v in enumerate(
                    self.display_df_a[key_column]
                )
                if str(v).strip()
            }

            right_key_map = {
                normalize_key(v): idx
                for idx, v in enumerate(
                    self.display_df_b[key_column]
                )
                if str(v).strip()
            }

            left_col_map = {
                col: idx
                for idx, col in enumerate(
                    self.df_a.columns
                )
            }

            right_col_map = {
                col: idx
                for idx, col in enumerate(
                    self.df_b.columns
                )
            }

            print("Left Keys:", left_key_map)
            print("Right Keys:", right_key_map)

            for _, diff in self.differences.iterrows():

                status = str(diff["Status"]).upper()
                key = normalize_key(diff["Key"])

                print(
                    "Processing:",
                    status,
                    key
                )

                # ==================================
                # MODIFIED CELL
                # ==================================

                if status == "MODIFIED":

                    column = diff["Column"]

                    if (
                        key in left_key_map
                        and key in right_key_map
                        and column in left_col_map
                        and column in right_col_map
                    ):

                        left_row = left_key_map[key]
                        right_row = right_key_map[key]

                        left_col = left_col_map[column]
                        right_col = right_col_map[column]

                        # Highlight entire row gray first

                        for col_idx in range(len(self.df_a.columns)):

                            self.left_sheet.highlight_cells(
                                row=left_row,
                                column=col_idx,
                                bg="#FFF59D"
                            )

                        for col_idx in range(len(self.df_b.columns)):

                            self.right_sheet.highlight_cells(
                                row=right_row,
                                column=col_idx,
                                bg="#FFF59D"
                            )

                        # Then highlight changed cell yellow

                        self.left_sheet.highlight_cells(
                            row=left_row,
                            column=left_col,
                            bg="#FFE817"
                        )

                        self.right_sheet.highlight_cells(
                            row=right_row,
                            column=right_col,
                            bg="#FFE817"
                        )

                # ==================================
                # DELETED ROW
                # ==================================

                elif status == "DELETED":

                    if key in left_key_map:

                        row_idx = left_key_map[key]

                        for col_idx in range(
                            len(self.df_a.columns)
                        ):

                            self.left_sheet.highlight_cells(
                                row=row_idx,
                                column=col_idx,
                                bg="#FF9999"
                            )

                # ==================================
                # ADDED ROW
                # ==================================

                elif status == "ADDED":

                    if key in right_key_map:

                        row_idx = right_key_map[key]

                        for col_idx in range(
                            len(self.df_b.columns)
                        ):

                            self.right_sheet.highlight_cells(
                                row=row_idx,
                                column=col_idx,
                                bg="#FF9999"
                            )

            # Force redraw
            self.left_sheet.refresh()
            self.right_sheet.refresh()

        except Exception as e:

            print("Highlight Error:", e)

    def export_report(self):

        if self.differences.empty:
            messagebox.showwarning(
                "Warning",
                "No comparison results available."
            )
            return

        save_path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel Files", "*.xlsx")]
        )

        if save_path:

            self.differences.to_excel(
                save_path,
                index=False
            )

            messagebox.showinfo(
                "Success",
                "Report exported successfully."
            )

    def clear_all(self):

        self.left_sheet.set_sheet_data([])
        self.right_sheet.set_sheet_data([])

        self.file_a = ""
        self.file_b = ""

        self.df_a = None
        self.df_b = None

        self.differences = pd.DataFrame()

        self.file_a_entry.delete(0, "end")
        self.file_b_entry.delete(0, "end")

        self.status_label.configure(
            text="Ready"
        )

    def toggle_theme(self):

        if self.theme_switch.get():
            ctk.set_appearance_mode("dark")
        else:
            ctk.set_appearance_mode("light")

    def clear_highlights(self):
        """
        Remove all cell highlights from both sheets.
        """

        try:

            self.left_sheet.dehighlight_all()
            self.right_sheet.dehighlight_all()

            self.left_sheet.refresh()
            self.right_sheet.refresh()

        except Exception as e:
            print("Clear Highlight Error:", e)


if __name__ == "__main__":

    ctk.set_appearance_mode("dark")
    ctk.set_default_color_theme("blue")

    import ctypes

    myappid = "sheetdiff.v1.0"
    ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(
        myappid
    )

    root = ctk.CTk()
    root.iconbitmap("SheetDiffLogo.ico")
    root.state("zoomed")



    app = ExcelComparisonTool(root)

    root.mainloop()
