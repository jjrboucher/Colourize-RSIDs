import tkinter as tk
from tkinter import ttk, filedialog, colorchooser, messagebox, scrolledtext
import zipfile
import xml.etree.ElementTree as ET
from xml.dom import minidom
import re
import os
import tempfile
import shutil
from collections import defaultdict


class WordRsidColorizer:
    def __init__(self, root):
        self.root = root
        self.root.title("Word OOXML rsidR Colorizer")
        self.root.geometry("1200x800")

        # Variables
        self.current_file = None
        self.rsid_colors = {}
        self.rsid_buttons = {}
        self.document_xml = None
        self.document_tree = None
        self.temp_dir = None

        # XML namespaces for Word documents
        self.namespaces = {
            'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
        }

        self.setup_ui()

    def setup_ui(self):
        # Main frame
        main_frame = ttk.Frame(self.root)
        main_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        # Top frame for file selection
        top_frame = ttk.Frame(main_frame)
        top_frame.pack(fill=tk.X, pady=(0, 10))

        ttk.Button(top_frame, text="Select Word Document",
                   command=self.select_file).pack(side=tk.LEFT)

        self.file_label = ttk.Label(top_frame, text="No file selected")
        self.file_label.pack(side=tk.LEFT, padx=(10, 0))

        ttk.Button(top_frame, text="Save Document",
                   command=self.save_document).pack(side=tk.RIGHT)

        # Content frame
        content_frame = ttk.Frame(main_frame)
        content_frame.pack(fill=tk.BOTH, expand=True)

        # Left panel for rsidR buttons
        left_panel = ttk.LabelFrame(content_frame, text="rsidR Values", width=300)
        left_panel.pack(side=tk.LEFT, fill=tk.Y, padx=(0, 10))
        left_panel.pack_propagate(False)

        # Create a frame to hold the canvas and scrollbar
        scroll_frame = ttk.Frame(left_panel)
        scroll_frame.pack(fill=tk.BOTH, expand=True)

        # Canvas with scrollbar
        self.canvas = tk.Canvas(scroll_frame, highlightthickness=0)
        self.scrollbar = ttk.Scrollbar(scroll_frame, orient="vertical", command=self.canvas.yview)
        self.scrollable_frame = ttk.Frame(self.canvas)

        # Configure scrolling
        self.scrollable_frame.bind(
            "<Configure>",
            lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all"))
        )

        # Create window in canvas
        self.canvas_window = self.canvas.create_window((0, 0), window=self.scrollable_frame, anchor="nw")
        self.canvas.configure(yscrollcommand=self.scrollbar.set)

        # Pack canvas and scrollbar
        self.canvas.pack(side="left", fill="both", expand=True)
        self.scrollbar.pack(side="right", fill="y")

        # Bind mouse wheel to canvas
        self.canvas.bind("<MouseWheel>", self._on_mousewheel)
        self.canvas.bind("<Button-4>", self._on_mousewheel)
        self.canvas.bind("<Button-5>", self._on_mousewheel)

        # Bind canvas resize to adjust scrollable frame width
        self.canvas.bind("<Configure>", self._on_canvas_configure)

        # Right panel for document preview
        right_panel = ttk.LabelFrame(content_frame, text="Document Preview")
        right_panel.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True)

        self.preview_text = scrolledtext.ScrolledText(right_panel, wrap=tk.WORD,
                                                      state=tk.DISABLED, font=("Arial", 10))
        self.preview_text.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)

    def _on_mousewheel(self, event):
        """Handle mouse wheel scrolling"""
        if event.delta:
            # Windows
            self.canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")
        else:
            # Linux
            if event.num == 4:
                self.canvas.yview_scroll(-1, "units")
            elif event.num == 5:
                self.canvas.yview_scroll(1, "units")

    def _on_canvas_configure(self, event):
        """Adjust the scrollable frame width when canvas is resized"""
        canvas_width = event.width
        self.canvas.itemconfig(self.canvas_window, width=canvas_width)

    def select_file(self):
        file_path = filedialog.askopenfilename(
            title="Select Word Document",
            filetypes=[("Word Documents", "*.docx"), ("All Files", "*.*")]
        )

        if file_path:
            self.current_file = file_path
            self.file_label.config(text=os.path.basename(file_path))
            self.load_document()

    def load_document(self):
        try:
            # Create temporary directory
            if self.temp_dir:
                shutil.rmtree(self.temp_dir, ignore_errors=True)
            self.temp_dir = tempfile.mkdtemp()

            # Extract docx file
            with zipfile.ZipFile(self.current_file, 'r') as zip_ref:
                zip_ref.extractall(self.temp_dir)

            # Load settings.xml to get rsidR values
            settings_path = os.path.join(self.temp_dir, 'word', 'settings.xml')
            rsid_values = self.extract_rsid_values(settings_path)

            # Load document.xml
            doc_path = os.path.join(self.temp_dir, 'word', 'document.xml')
            self.load_document_xml(doc_path)

            # Create buttons for rsidR values
            self.create_rsid_buttons(rsid_values)

            # Update preview
            self.update_preview()

        except Exception as e:
            messagebox.showerror("Error", f"Failed to load document: {str(e)}")

    def extract_rsid_values(self, settings_path):
        rsid_values = set()

        try:
            tree = ET.parse(settings_path)
            root = tree.getroot()

            print(f"Root element: {root.tag}")
            print(f"Root namespaces: {root.attrib}")

            # Find all rsid-related elements in settings.xml
            # Look for rsids container
            rsids_elem = root.find('.//w:rsids', self.namespaces)
            if rsids_elem is not None:
                print(f"Found rsids element with {len(list(rsids_elem))} children")

                # Get rsidRoot
                rsid_root = rsids_elem.find('./w:rsidRoot', self.namespaces)
                if rsid_root is not None:
                    val = rsid_root.get(f'{{{self.namespaces["w"]}}}val')
                    if val:
                        rsid_values.add(val)
                        print(f"Added rsidRoot: {val}")

                # Get all generic rsid values (this is what your document has)
                for rsid_elem in rsids_elem.findall('./w:rsid', self.namespaces):
                    val = rsid_elem.get(f'{{{self.namespaces["w"]}}}val')
                    if val:
                        rsid_values.add(val)
                        print(f"Added rsid: {val}")

                # Also check for other specific rsid types
                for rsid_r in rsids_elem.findall('./w:rsidR', self.namespaces):
                    val = rsid_r.get(f'{{{self.namespaces["w"]}}}val')
                    if val:
                        rsid_values.add(val)
                        print(f"Added rsidR: {val}")

                for rsid_rpr in rsids_elem.findall('./w:rsidRPr', self.namespaces):
                    val = rsid_rpr.get(f'{{{self.namespaces["w"]}}}val')
                    if val:
                        rsid_values.add(val)
                        print(f"Added rsidRPr: {val}")

                for rsid_del in rsids_elem.findall('./w:rsidDel', self.namespaces):
                    val = rsid_del.get(f'{{{self.namespaces["w"]}}}val')
                    if val:
                        rsid_values.add(val)
                        print(f"Added rsidDel: {val}")

                for rsid_p in rsids_elem.findall('./w:rsidP', self.namespaces):
                    val = rsid_p.get(f'{{{self.namespaces["w"]}}}val')
                    if val:
                        rsid_values.add(val)
                        print(f"Added rsidP: {val}")
            else:
                print("No rsids element found")
                # Debug: print all elements to see what's available
                for elem in root.iter():
                    print(f"Element: {elem.tag}, attrib: {elem.attrib}")

            print(f"Total found {len(rsid_values)} rsid values: {sorted(list(rsid_values))}")

        except Exception as e:
            print(f"Error extracting rsidR values: {e}")
            import traceback
            traceback.print_exc()

        return sorted(list(rsid_values))

    def load_document_xml(self, doc_path):
        try:
            with open(doc_path, 'r', encoding='utf-8') as f:
                self.document_xml = f.read()

            self.document_tree = ET.parse(doc_path)

        except Exception as e:
            print(f"Error loading document XML: {e}")

    def count_rsid_usage(self, rsid_values):
        """Count how many times each rsidR value appears in the document"""
        rsid_counts = {rsid: 0 for rsid in rsid_values}

        try:
            if not self.document_tree:
                return rsid_counts

            root = self.document_tree.getroot()

            # Count rsidR attributes in all elements
            for elem in root.iter():
                # Check for rsidR attribute
                rsid_r = elem.get(f'{{{self.namespaces["w"]}}}rsidR')
                if rsid_r and rsid_r in rsid_counts:
                    rsid_counts[rsid_r] += 1

                # Check for rsidP attribute (paragraph rsid)
                rsid_p = elem.get(f'{{{self.namespaces["w"]}}}rsidP')
                if rsid_p and rsid_p in rsid_counts:
                    rsid_counts[rsid_p] += 1

                # Check for rsidRPr attribute (run properties rsid)
                rsid_rpr = elem.get(f'{{{self.namespaces["w"]}}}rsidRPr')
                if rsid_rpr and rsid_rpr in rsid_counts:
                    rsid_counts[rsid_rpr] += 1

                # Check for rsidDel attribute (deletion rsid)
                rsid_del = elem.get(f'{{{self.namespaces["w"]}}}rsidDel')
                if rsid_del and rsid_del in rsid_counts:
                    rsid_counts[rsid_del] += 1

            print(f"rsidR usage counts: {rsid_counts}")

        except Exception as e:
            print(f"Error counting rsid usage: {e}")
            import traceback
            traceback.print_exc()

        return rsid_counts

    def create_rsid_buttons(self, rsid_values):
        # Clear existing buttons
        for widget in self.scrollable_frame.winfo_children():
            widget.destroy()

        self.rsid_buttons.clear()

        # Count rsid usage in document
        rsid_counts = self.count_rsid_usage(rsid_values)

        # Filter out rsidR values with zero count
        active_rsids = [(rsid, count) for rsid, count in rsid_counts.items() if count > 0]

        # Create buttons for each active rsidR value
        if not active_rsids:
            no_rsid_label = ttk.Label(self.scrollable_frame,
                                      text="No active rsidR values found in document")
            no_rsid_label.pack(pady=10)
            return

        # Sort by count (descending) for better UX
        active_rsids.sort(key=lambda x: x[1], reverse=True)

        for rsid, count in active_rsids:
            button_frame = ttk.Frame(self.scrollable_frame)
            button_frame.pack(fill=tk.X, pady=2, padx=5)

            button_text = f"rsid: {rsid} ({count})"

            button = tk.Button(
                button_frame,
                text=button_text,
                width=25,
                height=2,
                bg='white',
                font=("Arial", 8),
                command=lambda r=rsid: self.select_color(r)
            )
            button.pack(side=tk.LEFT, padx=2)

            # Reset button
            reset_btn = ttk.Button(
                button_frame,
                text="Reset",
                width=8,
                command=lambda r=rsid: self.reset_color(r)
            )
            reset_btn.pack(side=tk.RIGHT, padx=2)

            self.rsid_buttons[rsid] = button

        # Update scroll region after adding buttons
        self.scrollable_frame.update_idletasks()
        self.canvas.configure(scrollregion=self.canvas.bbox("all"))

    def select_color(self, rsid):
        color = colorchooser.askcolor(title=f"Choose color for {rsid}")
        if color[1]:  # If a color was selected
            hex_color = color[1]
            self.rsid_colors[rsid] = hex_color
            self.rsid_buttons[rsid].config(bg=hex_color)
            self.update_preview()

    def reset_color(self, rsid):
        if rsid in self.rsid_colors:
            del self.rsid_colors[rsid]
        self.rsid_buttons[rsid].config(bg='white')
        self.update_preview()

    def update_preview(self):
        if not self.document_tree:
            return

        try:
            # Extract text with rsidR information
            text_content = self.extract_text_with_rsid()

            # Update preview text widget
            self.preview_text.config(state=tk.NORMAL)
            self.preview_text.delete(1.0, tk.END)

            # Configure tags for colors
            for rsid, color in self.rsid_colors.items():
                self.preview_text.tag_configure(f"rsid_{rsid}",
                                                foreground=color,
                                                font=("Arial", 10, "bold"))

            # Configure default tag
            self.preview_text.tag_configure("default", font=("Arial", 10))

            # Insert text with appropriate tags
            for text, rsid in text_content:
                if rsid and rsid in self.rsid_colors:
                    self.preview_text.insert(tk.END, text, f"rsid_{rsid}")
                else:
                    self.preview_text.insert(tk.END, text, "default")

            # Add debug information
            if text_content:
                self.preview_text.insert(tk.END, f"\n\n--- Debug Info ---\n")
                self.preview_text.insert(tk.END, f"Total text segments: {len(text_content)}\n")
                rsids_found = set(rsid for _, rsid in text_content if rsid)
                self.preview_text.insert(tk.END, f"rsidR values found in text: {sorted(list(rsids_found))}\n")
                self.preview_text.insert(tk.END, f"Available colors: {list(self.rsid_colors.keys())}\n")

            self.preview_text.config(state=tk.DISABLED)

        except Exception as e:
            print(f"Error updating preview: {e}")
            import traceback
            traceback.print_exc()

    def extract_text_with_rsid(self):
        text_content = []

        try:
            root = self.document_tree.getroot()

            # Process paragraphs to maintain document structure
            for para in root.findall('.//w:p', self.namespaces):
                para_rsid = para.get(f'{{{self.namespaces["w"]}}}rsidR')
                para_rsid_p = para.get(f'{{{self.namespaces["w"]}}}rsidP')

                # Process runs within paragraph
                para_text = []
                for run in para.findall('.//w:r', self.namespaces):
                    run_rsid = run.get(f'{{{self.namespaces["w"]}}}rsidR')

                    # Use run rsid, then paragraph rsid as fallback
                    effective_rsid = run_rsid or para_rsid or para_rsid_p

                    # Extract text from this run
                    text_parts = []
                    for t_elem in run.findall('./w:t', self.namespaces):
                        if t_elem.text:
                            text_parts.append(t_elem.text)

                    # Handle tabs
                    for tab in run.findall('./w:tab', self.namespaces):
                        text_parts.append('\t')

                    # Handle breaks
                    for br in run.findall('./w:br', self.namespaces):
                        text_parts.append('\n')

                    if text_parts:
                        text = ''.join(text_parts)
                        para_text.append((text, effective_rsid))

                # Add paragraph content
                if para_text:
                    text_content.extend(para_text)
                    text_content.append(('\n\n', None))  # Paragraph break
                elif para_rsid or para_rsid_p:
                    # Empty paragraph but has rsid
                    text_content.append(('\n', para_rsid or para_rsid_p))

        except Exception as e:
            print(f"Error extracting text with rsid: {e}")
            import traceback
            traceback.print_exc()

        return text_content

    def save_document(self):
        if not self.current_file or not self.document_tree:
            messagebox.showwarning("Warning", "No document loaded")
            return

        save_path = filedialog.asksaveasfilename(
            title="Save Colored Document",
            defaultextension=".docx",
            filetypes=[("Word Documents", "*.docx"), ("All Files", "*.*")]
        )

        if save_path:
            try:
                self.apply_colors_to_document()

                # Create new docx file
                with zipfile.ZipFile(save_path, 'w', zipfile.ZIP_DEFLATED) as zip_out:
                    for root, dirs, files in os.walk(self.temp_dir):
                        for file in files:
                            file_path = os.path.join(root, file)
                            arc_path = os.path.relpath(file_path, self.temp_dir)
                            zip_out.write(file_path, arc_path)

                messagebox.showinfo("Success", f"Document saved to {save_path}")

            except Exception as e:
                messagebox.showerror("Error", f"Failed to save document: {str(e)}")

    def apply_colors_to_document(self):
        if not self.document_tree or not self.rsid_colors:
            return

        try:
            root = self.document_tree.getroot()

            # Apply colors to runs with matching rsidR values
            for run in root.findall('.//w:r', self.namespaces):
                rsid = run.get(f'{{{self.namespaces["w"]}}}rsidR')

                if rsid in self.rsid_colors:
                    hex_color = self.rsid_colors[rsid].lstrip('#').upper()

                    # Find or create run properties
                    rpr = run.find('./w:rPr', self.namespaces)
                    if rpr is None:
                        rpr = ET.SubElement(run, f'{{{self.namespaces["w"]}}}rPr')
                        # Insert rPr as first child
                        run.insert(0, rpr)

                    # Remove existing color element if present
                    existing_color = rpr.find('./w:color', self.namespaces)
                    if existing_color is not None:
                        rpr.remove(existing_color)

                    # Add new color element
                    color_elem = ET.SubElement(rpr, f'{{{self.namespaces["w"]}}}color')
                    color_elem.set(f'{{{self.namespaces["w"]}}}val', hex_color)

            # Also apply colors to paragraphs with matching rsidP values
            for para in root.findall('.//w:p', self.namespaces):
                rsid_p = para.get(f'{{{self.namespaces["w"]}}}rsidP')

                if rsid_p in self.rsid_colors:
                    hex_color = self.rsid_colors[rsid_p].lstrip('#').upper()

                    # Apply color to all runs in this paragraph that don't already have color
                    for run in para.findall('.//w:r', self.namespaces):
                        # Skip if run already has rsidR with color applied
                        run_rsid = run.get(f'{{{self.namespaces["w"]}}}rsidR')
                        if run_rsid and run_rsid in self.rsid_colors:
                            continue

                        rpr = run.find('./w:rPr', self.namespaces)
                        if rpr is None:
                            rpr = ET.SubElement(run, f'{{{self.namespaces["w"]}}}rPr')
                            run.insert(0, rpr)

                        # Check if color already exists
                        existing_color = rpr.find('./w:color', self.namespaces)
                        if existing_color is None:
                            color_elem = ET.SubElement(rpr, f'{{{self.namespaces["w"]}}}color')
                            color_elem.set(f'{{{self.namespaces["w"]}}}val', hex_color)

            # Save modified document.xml with proper XML formatting
            doc_path = os.path.join(self.temp_dir, 'word', 'document.xml')

            # Register namespace to avoid ns0: prefixes
            ET.register_namespace('w', self.namespaces['w'])

            # Save the XML with proper formatting
            with open(doc_path, 'wb') as f:
                # Write XML declaration
                f.write(b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n')

                # Write the tree without additional XML declaration
                tree_bytes = ET.tostring(root, encoding='utf-8', xml_declaration=False)
                f.write(tree_bytes)

        except Exception as e:
            print(f"Error applying colors to document: {e}")
            import traceback
            traceback.print_exc()

    def __del__(self):
        if hasattr(self, 'temp_dir') and self.temp_dir:
            shutil.rmtree(self.temp_dir, ignore_errors=True)


def main():
    root = tk.Tk()
    app = WordRsidColorizer(root)
    root.mainloop()


if __name__ == "__main__":
    main()