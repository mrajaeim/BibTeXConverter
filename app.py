import tkinter as tk
from tkinter import filedialog, messagebox
from tkinter import ttk
import xml.etree.ElementTree as ET
import bibtexparser
import os

class BibTeXConverter:
    def __init__(self, root):
        self.root = root
        self.root.title("BibTeX to Word XML Converter")
        self.root.geometry("600x300")
        
        # Configure style
        style = ttk.Style()
        style.configure('TButton', padding=5)
        style.configure('TLabel', padding=5)
        
        # Create main frame
        self.main_frame = ttk.Frame(root, padding="20")
        self.main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Create and configure grid
        self.main_frame.columnconfigure(0, weight=1)
        self.main_frame.rowconfigure(3, weight=1)
        
        # Components
        self.create_widgets()
        
    def create_widgets(self):
        # Header
        header_label = ttk.Label(
            self.main_frame,
            text="Convert BibTeX files to Word Bibliography XML",
            font=('Helvetica', 12, 'bold')
        )
        header_label.grid(row=0, column=0, columnspan=2, pady=10)
        
        # Input file section
        self.input_frame = ttk.Frame(self.main_frame)
        self.input_frame.grid(row=1, column=0, sticky=(tk.W, tk.E), pady=5)
        
        self.input_label = ttk.Label(self.input_frame, text="Selected file: None")
        self.input_label.grid(row=0, column=0, sticky=tk.W)
        
        self.select_button = ttk.Button(
            self.input_frame,
            text="Select BibTeX File",
            command=self.select_file
        )
        self.select_button.grid(row=0, column=1, padx=5)
        
        # Convert button
        self.convert_button = ttk.Button(
            self.main_frame,
            text="Convert to XML",
            command=self.convert_file,
            state=tk.DISABLED
        )
        self.convert_button.grid(row=2, column=0, pady=20)
        
        # Status label
        self.status_label = ttk.Label(
            self.main_frame,
            text="Ready",
            font=('Helvetica', 10)
        )
        self.status_label.grid(row=3, column=0, sticky=tk.W)
        
        self.input_path = None
        
    def select_file(self):
        try:
            filename = filedialog.askopenfilename(
                filetypes=[("BibTeX files", "*.bib"), ("All files", "*.*")]
            )
            if filename:
                self.input_path = filename
                self.input_label.config(text=f"Selected file: {os.path.basename(filename)}")
                self.convert_button.config(state=tk.NORMAL)
                self.status_label.config(text="File selected. Ready to convert.")
        except Exception as e:
            messagebox.showerror("Error", f"Error selecting file: {str(e)}")
            
    def convert_file(self):
        if not self.input_path:
            messagebox.showerror("Error", "Please select a BibTeX file first.")
            return
            
        try:
            # Read BibTeX file
            with open(self.input_path, 'r', encoding='utf-8') as bibtex_file:
                bib_database = bibtexparser.load(bibtex_file)
            
            # Create XML structure
            sources = ET.Element('Sources')
            sources.set('xmlns', 'http://schemas.openxmlformats.org/officeDocument/2006/bibliography')
            
            for entry in bib_database.entries:
                source = ET.SubElement(sources, 'Source')
                
                # Tag
                tag = ET.SubElement(source, 'Tag')
                tag.text = entry.get('ID', '')
                
                # Type mapping
                type_mapping = {
                    'article': 'JournalArticle',
                    'book': 'Book',
                    'inproceedings': 'ConferenceProceedings',
                    'thesis': 'Report',
                    'phdthesis': 'Report',
                    'mastersthesis': 'Report'
                }
                
                source_type = ET.SubElement(source, 'SourceType')
                source_type.text = type_mapping.get(entry.get('ENTRYTYPE', '').lower(), 'Book')
                
                # Authors
                if 'author' in entry:
                    authors = ET.SubElement(source, 'Author')
                    author_list = entry['author'].split(' and ')
                    for author in author_list:
                        author_element = ET.SubElement(authors, 'Author')
                        namelist = ET.SubElement(author_element, 'NameList')
                        
                        # Simple name parsing
                        name_parts = author.strip().split(',')
                        if len(name_parts) > 1:
                            last = ET.SubElement(namelist, 'Last')
                            last.text = name_parts[0].strip()
                            first = ET.SubElement(namelist, 'First')
                            first.text = name_parts[1].strip()
                        else:
                            name_parts = author.strip().split()
                            if name_parts:
                                last = ET.SubElement(namelist, 'Last')
                                last.text = name_parts[-1]
                                if len(name_parts) > 1:
                                    first = ET.SubElement(namelist, 'First')
                                    first.text = ' '.join(name_parts[:-1])
                
                # Title
                if 'title' in entry:
                    title = ET.SubElement(source, 'Title')
                    title.text = entry['title']
                
                # Year
                if 'year' in entry:
                    year = ET.SubElement(source, 'Year')
                    year.text = entry['year']
                
                # Journal/Publisher
                if 'journal' in entry:
                    journal = ET.SubElement(source, 'JournalName')
                    journal.text = entry['journal']
                elif 'publisher' in entry:
                    publisher = ET.SubElement(source, 'Publisher')
                    publisher.text = entry['publisher']
            
            # Save dialog
            output_path = filedialog.asksaveasfilename(
                defaultextension=".xml",
                filetypes=[("XML files", "*.xml"), ("All files", "*.*")]
            )
            
            if output_path:
                tree = ET.ElementTree(sources)
                tree.write(output_path, encoding='utf-8', xml_declaration=True)
                self.status_label.config(text="Conversion completed successfully!")
                messagebox.showinfo("Success", "File converted successfully!")
            
        except Exception as e:
            messagebox.showerror("Error", f"Error during conversion: {str(e)}")
            self.status_label.config(text="Conversion failed. See error message.")

def main():
    root = tk.Tk()
    app = BibTeXConverter(root)
    root.mainloop()

if __name__ == "__main__":
    main()