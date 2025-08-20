#!/usr/bin/env python3
"""
Address Label Application
Creates printable address labels with preview and sizing options
"""

import os
import tkinter as tk
from tkinter import font as tkFont
from tkinter import ttk, messagebox, filedialog


class AddressLabelApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Address Label Maker")
        self.root.geometry("800x600")

        # Return address (constant)
        self.return_address = """Bruce Eckel
107 WhiteRock Suite 969
Crested Butte, CO 81224"""

        # Variables
        self.to_address = tk.StringVar()
        self.label_size = tk.DoubleVar(value=1.0)  # Scale factor
        self.indent_amount = tk.DoubleVar(value=3.0)  # Indentation in inches

        self.create_widgets()
        self.update_preview()

    def create_widgets(self):
        # Main frame
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))

        # Configure grid weights
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(1, weight=1)
        main_frame.rowconfigure(2, weight=1)

        # To Address input
        ttk.Label(main_frame, text="To Address:").grid(row=0, column=0, sticky=tk.W, pady=(0, 5))

        # Text widget for multi-line address input
        address_frame = ttk.Frame(main_frame)
        address_frame.grid(row=0, column=1, sticky=(tk.W, tk.E), pady=(0, 10))
        address_frame.columnconfigure(0, weight=1)

        self.address_text = tk.Text(address_frame, height=4, width=40, wrap=tk.WORD)
        scrollbar = ttk.Scrollbar(address_frame, orient="vertical", command=self.address_text.yview)
        self.address_text.configure(yscrollcommand=scrollbar.set)

        self.address_text.grid(row=0, column=0, sticky=(tk.W, tk.E))
        scrollbar.grid(row=0, column=1, sticky=(tk.N, tk.S))

        # Bind text changes to preview update
        self.address_text.bind('<KeyRelease>', self.on_address_change)
        self.address_text.bind('<Button-1>', self.on_address_change)

        # Size and indent controls
        controls_frame = ttk.Frame(main_frame)
        controls_frame.grid(row=1, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=10)

        # Size control
        size_frame = ttk.Frame(controls_frame)
        size_frame.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 20))

        ttk.Label(size_frame, text="Label Size:").pack(side=tk.LEFT)

        size_scale = ttk.Scale(size_frame, from_=0.5, to=3.0,
                               orient=tk.HORIZONTAL, variable=self.label_size,
                               command=self.on_size_change, length=150)
        size_scale.pack(side=tk.LEFT, padx=(10, 10))

        self.size_label = ttk.Label(size_frame, text="1.0x")
        self.size_label.pack(side=tk.LEFT)

        # Indent control
        indent_frame = ttk.Frame(controls_frame)
        indent_frame.pack(side=tk.LEFT, fill=tk.X, expand=True)

        ttk.Label(indent_frame, text="Indent:").pack(side=tk.LEFT)

        indent_scale = ttk.Scale(indent_frame, from_=0.5, to=6.0,
                                orient=tk.HORIZONTAL, variable=self.indent_amount,
                                command=self.on_indent_change, length=150)
        indent_scale.pack(side=tk.LEFT, padx=(10, 10))

        self.indent_label = ttk.Label(indent_frame, text="3.0\"")
        self.indent_label.pack(side=tk.LEFT)

        # Preview canvas
        preview_frame = ttk.LabelFrame(main_frame, text="Preview", padding="5")
        preview_frame.grid(row=2, column=0, columnspan=2, sticky=(tk.W, tk.E, tk.N, tk.S), pady=10)
        preview_frame.columnconfigure(0, weight=1)
        preview_frame.rowconfigure(0, weight=1)

        # Canvas with scrollbars
        canvas_frame = ttk.Frame(preview_frame)
        canvas_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        canvas_frame.columnconfigure(0, weight=1)
        canvas_frame.rowconfigure(0, weight=1)

        self.canvas = tk.Canvas(canvas_frame, bg='white', width=700, height=400)
        v_scrollbar = ttk.Scrollbar(canvas_frame, orient="vertical", command=self.canvas.yview)
        h_scrollbar = ttk.Scrollbar(canvas_frame, orient="horizontal", command=self.canvas.xview)
        self.canvas.configure(yscrollcommand=v_scrollbar.set, xscrollcommand=h_scrollbar.set)

        self.canvas.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        v_scrollbar.grid(row=0, column=1, sticky=(tk.N, tk.S))
        h_scrollbar.grid(row=1, column=0, sticky=(tk.W, tk.E))

        # Buttons
        button_frame = ttk.Frame(main_frame)
        button_frame.grid(row=3, column=0, columnspan=2, pady=10)

        ttk.Button(button_frame, text="Clear", command=self.clear_address).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="Save as PDF", command=self.save_pdf).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="Print", command=self.print_label).pack(side=tk.LEFT, padx=5)

    def on_address_change(self, event=None):
        """Update preview when address text changes"""
        self.root.after_idle(self.update_preview)

    def on_size_change(self, value):
        """Update preview when size changes"""
        self.size_label.config(text=f"{float(value):.1f}x")
        self.update_preview()

    def on_indent_change(self, value):
        """Update preview when indent changes"""
        self.indent_label.config(text=f"{float(value):.1f}\"")
        self.update_preview()

    def clear_address(self):
        """Clear the to address field"""
        self.address_text.delete(1.0, tk.END)
        self.update_preview()

    def update_preview(self):
        """Update the label preview"""
        self.canvas.delete("all")

        # Get current to address
        to_addr = self.address_text.get(1.0, tk.END).strip()

        # Calculate dimensions (11" x 8.5" at 72 DPI for landscape, scaled down for preview)
        base_width = 792  # 11 * 72 (landscape)
        base_height = 612  # 8.5 * 72 (landscape)
        scale_factor = 0.6  # Scale down for preview
        preview_width = int(base_width * scale_factor)
        preview_height = int(base_height * scale_factor)

        # Update canvas scroll region
        self.canvas.configure(scrollregion=(0, 0, preview_width, preview_height))

        # Draw page border
        self.canvas.create_rectangle(5, 5, preview_width - 5, preview_height - 5,
                                     outline='lightgray', width=2)

        # Calculate font sizes based on label size
        size_multiplier = self.label_size.get()
        return_font_size = int(12 * scale_factor * size_multiplier)
        to_font_size = int(16 * scale_factor * size_multiplier)

        # Create fonts
        return_font = tkFont.Font(family="Arial", size=return_font_size)
        to_font = tkFont.Font(family="Arial", size=to_font_size, weight="bold")

        # Calculate positions
        margin = 30 * scale_factor * size_multiplier
        center_x = preview_width // 2
        
        # Calculate indent position (convert inches to pixels)
        indent_pixels = self.indent_amount.get() * 72 * scale_factor * size_multiplier

        # Draw return address (top left)
        return_lines = self.return_address.split('\n')
        y_pos = margin
        for line in return_lines:
            self.canvas.create_text(margin, y_pos, text=line, font=return_font,
                                    anchor="nw", fill="black")
            y_pos += return_font_size + 2

        # Draw to address (left-justified with indentation)
        if to_addr:
            to_lines = to_addr.split('\n')
            # Calculate starting y position to center the address block vertically
            total_height = len(to_lines) * (to_font_size + 4)
            start_y = (preview_height - total_height) // 2

            y_pos = start_y
            for line in to_lines:
                if line.strip():  # Only draw non-empty lines
                    self.canvas.create_text(indent_pixels, y_pos, text=line.strip(),
                                            font=to_font, anchor="nw", fill="black")
                y_pos += to_font_size + 4
        else:
            # Show placeholder text
            self.canvas.create_text(indent_pixels, preview_height // 2,
                                    text="Enter destination address above",
                                    font=to_font, anchor="nw", fill="lightgray")

    def save_pdf(self):
        """Save label as PDF (requires reportlab)"""
        try:
            from reportlab.pdfgen import canvas
            from reportlab.lib.pagesizes import letter
            from reportlab.lib.units import inch
        except ImportError:
            messagebox.showerror("Error",
                                 "reportlab library not found.\n"
                                 "Install it with: pip install reportlab")
            return

        to_addr = self.address_text.get(1.0, tk.END).strip()
        if not to_addr:
            messagebox.showwarning("Warning", "Please enter a destination address.")
            return

        filename = filedialog.asksaveasfilename(
            defaultextension=".pdf",
            filetypes=[("PDF files", "*.pdf")],
            title="Save Label as PDF"
        )

        if filename:
            try:
                # Create PDF in landscape
                from reportlab.lib.pagesizes import landscape
                c = canvas.Canvas(filename, pagesize=landscape(letter))
                width, height = landscape(letter)

                size_multiplier = self.label_size.get()

                # Return address
                return_font_size = 12 * size_multiplier
                c.setFont("Helvetica", return_font_size)

                margin = 0.5 * inch
                y_pos = height - margin - return_font_size

                for line in self.return_address.split('\n'):
                    c.drawString(margin, y_pos, line)
                    y_pos -= return_font_size + 2

                # To address (left-justified with indentation)
                to_font_size = 16 * size_multiplier
                c.setFont("Helvetica-Bold", to_font_size)

                # Calculate indent position in points
                indent_points = self.indent_amount.get() * inch

                to_lines = [line.strip() for line in to_addr.split('\n') if line.strip()]
                total_height = len(to_lines) * (to_font_size + 4)
                start_y = (height + total_height) / 2

                for line in to_lines:
                    c.drawString(indent_points, start_y, line)
                    start_y -= to_font_size + 4

                c.save()
                messagebox.showinfo("Success", f"Label saved as {filename}")

            except Exception as e:
                messagebox.showerror("Error", f"Failed to save PDF: {str(e)}")

    def print_label(self):
        """Print the label"""
        to_addr = self.address_text.get(1.0, tk.END).strip()
        if not to_addr:
            messagebox.showwarning("Warning", "Please enter a destination address.")
            return

        # Create a simple HTML file for printing
        html_content = self.create_html_label(to_addr)

        # Save temporary HTML file
        temp_file = "temp_label.html"
        try:
            with open(temp_file, 'w', encoding='utf-8') as f:
                f.write(html_content)

            # Try to open with default browser for printing
            import webbrowser
            webbrowser.open(f"file://{os.path.abspath(temp_file)}")

            messagebox.showinfo("Print",
                                "Label opened in browser. Use Ctrl+P to print.\n"
                                "Make sure to select 'More settings' and choose appropriate paper size.")

        except Exception as e:
            messagebox.showerror("Error", f"Failed to create print file: {str(e)}")

    def create_html_label(self, to_addr):
        """Create HTML content for printing"""
        size_multiplier = self.label_size.get()
        return_font_size = int(12 * size_multiplier)
        to_font_size = int(16 * size_multiplier)

        indent_inches = self.indent_amount.get()
        
        return f"""
<!DOCTYPE html>
<html>
<head>
    <title>Address Label</title>
    <style>
        @page {{ size: letter landscape; margin: 0.5in; }}
        body {{ margin: 0; padding: 20px; font-family: Arial, sans-serif; }}
        .return-address {{ 
            position: absolute; 
            top: 30px; 
            left: 30px; 
            font-size: {return_font_size}px; 
            line-height: 1.2;
        }}
        .to-address {{ 
            position: absolute; 
            top: 50%; 
            left: {indent_inches}in; 
            transform: translateY(-50%); 
            font-size: {to_font_size}px; 
            font-weight: bold; 
            line-height: 1.3; 
            text-align: left;
            white-space: pre-line;
        }}
        @media print {{
            .no-print {{ display: none; }}
        }}
    </style>
</head>
<body>
    <div class="return-address">{self.return_address.replace(chr(10), '<br>')}</div>
    <div class="to-address">{to_addr}</div>
    <div class="no-print" style="position: fixed; bottom: 20px; right: 20px; background: #f0f0f0; padding: 10px; border-radius: 5px;">
        <p>Press Ctrl+P to print<br>Make sure paper size is set to Letter</p>
    </div>
</body>
</html>
"""


def main():
    root = tk.Tk()
    app = AddressLabelApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
