"""run the app"""

#! IMPORTS

import tkinter as tk
from os import makedirs
from os.path import dirname, join, sep, exists
from tkinter import filedialog, messagebox, ttk

from pandas import ExcelWriter
from labio.read.biostrength import PRODUCTS

#! BIOREADER CLASS


class Bioreader(tk.Tk):
    """generate the Bioreader GUI"""

    _label: tk.StringVar
    _product: tk.StringVar

    def open_file(self):
        """open file dialog"""
        file = filedialog.askopenfilename(
            parent=self,
            initialdir=join(dirname(__file__)),
            title="Please select a file:",
            filetypes=[("text files", ".txt")],
        )
        self._label.set(file)

    def save_file(self):
        """save file dialog"""
        try:
            # convert raw data into user-understandable data
            data = PRODUCTS[self._product.get()].from_file(self._label.get())
            data = data.as_dataframe()

            # check where to save the data
            initial = self._label.get().rsplit(".", 1)[0] + "_converted.xlsx"
            file = filedialog.asksaveasfilename(
                parent=self,
                confirmoverwrite=False,
                filetypes=[("Excel Files", ".xlsx")],
                initialfile=initial,
                title="Please select the file to be saved:",
            )
            file = file.replace("/", sep)

            # save the converted data
            if len(file) > 0:
                if not file.endswith(".xlsx"):
                    file += ".xlsx"
                makedirs(file.rsplit(sep, 1)[0], exist_ok=True)
                if exists(file):
                    exwriter = ExcelWriter(
                        file,
                        mode="a",
                        if_sheet_exists="replace",
                    )
                else:
                    exwriter = ExcelWriter(file, mode="w")
                with exwriter as wrt:
                    data.to_excel(excel_writer=wrt, sheet_name=self._product.get())
                messagebox.showinfo("Info", "Conversion complete")
        except Exception as exc:
            messagebox.showerror(
                title="Error",
                message=exc.args[0],
            )

    def show(self):
        """show the app"""
        self.mainloop()

    @property
    def icon(self):
        """pillow icon"""
        return join(dirname(__file__), "assets", "icon.png")

    def __init__(self):
        super().__init__()
        self.geometry("400x100")
        self.iconphoto(True, tk.PhotoImage(file=self.icon))
        self.title("Bioreader")

        # input frame
        input_frame = ttk.LabelFrame(master=self, text="Input data")
        input_frame.pack(side="top", fill="both", expand=True)
        self._label = tk.StringVar(value="")

        # input button
        button = ttk.Button(
            master=input_frame,
            text="Import file",
            command=self.open_file,
        )
        button.pack(side="left")
        label = ttk.Label(master=input_frame, textvariable=self._label)
        label.pack(side="right")

        # output frame
        output_frame = ttk.LabelFrame(master=self, text="Output data")
        output_frame.pack(side="bottom", fill="both", expand=True)
        pr_label = ttk.Label(master=output_frame, text="Product: ")
        pr_label.pack(side="left")
        products = list(PRODUCTS.keys())
        self._product = tk.StringVar(value=products[0])
        prods = [products[0]] + products
        dropdown = ttk.OptionMenu(output_frame, self._product, *prods)
        dropdown.pack(side="left")
        dropdown.config(padding=2)

        # save button
        save = ttk.Button(
            master=output_frame,
            text="Convert",
            command=self.save_file,
        )
        save.pack(side="right")


#! MAIN


if __name__ == "__main__":
    Bioreader().show()
