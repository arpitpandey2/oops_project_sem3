from PyPDF2 import PdfMerger, PdfReader, PdfWriter
import fitz  
from docx import Document

class PDFMerger:
    def __init__(self, pdf_list, output):
        self.pdf_list = pdf_list
        self.output = output
    
    def merge(self):
        merger = PdfMerger()
        for pdf in self.pdf_list:
            merger.append(pdf)
        merger.write(self.output)
        merger.close()
        print("PDFs merged successfully!")

class PDFReorder:
    def __init__(self, input_pdf, page_order, output_pdf):
        self.input_pdf = input_pdf
        self.page_order = page_order
        self.output_pdf = output_pdf
    
    def reorder(self):
        reader = PdfReader(self.input_pdf)
        writer = PdfWriter()
        
        for page_num in self.page_order:
            writer.add_page(reader.pages[page_num])
        
        with open(self.output_pdf, 'wb') as output_file:
            writer.write(output_file)
        print("PDF pages reordered successfully!")

class PDFCompressor:
    def __init__(self, input_pdf, output_pdf, zoom_x=0.6, zoom_y=0.6):
        self.input_pdf = input_pdf
        self.output_pdf = output_pdf
        self.zoom_x = zoom_x
        self.zoom_y = zoom_y
    
    def compress(self):
        pdf = fitz.open(self.input_pdf)
        mat = fitz.Matrix(self.zoom_x, self.zoom_y)  
        
        compressed_pdf = fitz.open()
        
        for page in pdf:
            pix = page.get_pixmap(matrix=mat, alpha=False)
            
            img_pdf = fitz.open()
            img_bytes = pix.tobytes()  
            rect = page.rect  
            compressed_page = compressed_pdf.new_page(width=rect.width, height=rect.height)
            compressed_page.insert_image(rect, stream=img_bytes)
        
        compressed_pdf.save(self.output_pdf)
        pdf.close()
        compressed_pdf.close()
        print("PDF compressed successfully!")

class DOCMerger:
    def __init__(self, doc_list, output):
        self.doc_list = doc_list
        self.output = output
    
    def merge(self):
        merged_doc = Document()
        
        for doc in self.doc_list:
            sub_doc = Document(doc)
            for element in sub_doc.element.body:
                merged_doc.element.body.append(element)
        
        merged_doc.save(self.output)
        print("DOC files merged successfully!")

class DocumentProcessor:
    def __init__(self):
        pass

    def process(self):
        print("Choose an operation:")
        print("1. Merge PDFs")
        print("2. Reorder PDF Pages")
        print("3. Compress PDF")
        print("4. Merge DOC Files")
        choice = int(input("Enter the number corresponding to your choice: "))
        
        if choice == 1:
            pdf_list = input("Enter PDF file names separated by commas: ").split(',')
            output = input("Enter the output file name (e.g., merged_output.pdf): ").strip()
            merger = PDFMerger([pdf.strip() for pdf in pdf_list], output)
            merger.merge()
        
        elif choice == 2:
            input_pdf = input("Enter the PDF file name to reorder: ").strip()
            page_order = list(map(int, input("Enter the new page order (e.g., 2,0,1): ").split(',')))
            output_pdf = input("Enter the output file name (e.g., reordered_output.pdf): ").strip()
            reorder = PDFReorder(input_pdf, page_order, output_pdf)
            reorder.reorder()
        
        elif choice == 3:
            input_pdf = input("Enter the PDF file name to compress: ").strip().strip('"').strip("'")
            output_pdf = input("Enter the output file name (e.g., compressed_output.pdf): ").strip().strip('"').strip("'")
            compressor = PDFCompressor(input_pdf, output_pdf)
            compressor.compress()
        
        elif choice == 4:
            doc_list = input("Enter DOC file names separated by commas: ").split(',')
            output = input("Enter the output file name (e.g., merged_output.docx): ").strip()
            doc_merger = DOCMerger([doc.strip() for doc in doc_list], output)
            doc_merger.merge()
        
        else:
            print("Invalid choice. Please select a valid option.")

if __name__ == "__main__":
    processor = DocumentProcessor()
    processor.process()
