import os
import sys
import fitz
import tempfile
import csv
import datetime
import PyPDF2
from PyPDF2 import PdfReader, PdfWriter
from pdf2image import convert_from_path
from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas
from subprocess import Popen, DEVNULL
from PyQt5.QtWidgets import QApplication, QMainWindow, QPushButton, QLabel, QVBoxLayout, QWidget, QFileDialog, QMessageBox, QProgressBar, QTreeWidget, QTreeWidgetItem
from PyQt5.QtWidgets import QGroupBox, QHBoxLayout


class PDFConverterApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.total_files_to_process = 0
        self.files_processed = 0
        self.processed_files_info = []
        self.log_tree = None
        self.converted_tree = QTreeWidget()
        self.corrupt_tree = None
        self.downloads = self.get_downloads_path()
        self.initUI()


    def initUI(self):
        self.setWindowTitle("PDF Converter")
        self.setGeometry(600, 600, 640, 640)


        # Create a group for the first two buttons
        button_group_1 = QGroupBox("Actions")
        scan_button = QPushButton("Scan F Drive")
        scan_button.clicked.connect(self.scan_f_drive)
        convert_button = QPushButton("Start Converting")
        convert_button.clicked.connect(self.start_converting)

        button_layout_1 = QVBoxLayout()
        button_layout_1.addWidget(scan_button)
        button_layout_1.addWidget(convert_button)
        button_group_1.setLayout(button_layout_1)

        # Create a group for the last two buttons
        button_group_2 = QGroupBox("Actions")
        reset_button = QPushButton("Reset")
        reset_button.clicked.connect(self.reset)
        exit_button = QPushButton("Exit")
        exit_button.clicked.connect(self.close)

        button_layout_2 = QVBoxLayout()
        button_layout_2.addWidget(reset_button)
        button_layout_2.addWidget(exit_button)
        button_group_2.setLayout(button_layout_2)

        self.result_label = QLabel()
        self.progress_bar = QProgressBar()

        # Create QTreeWidget to display converted files
        self.log_tree = QTreeWidget()
        self.log_tree.setHeaderLabels(["File Name", "Status", "Location"])
        self.log_tree.setColumnWidth(0, 200)
        self.log_tree.setColumnWidth(1, 200)

        self.corrupt_tree = QTreeWidget()
        self.corrupt_tree.setHeaderLabels(["File Name", "Status", "Location"])
        self.corrupt_tree.setColumnWidth(0, 200)
        self.corrupt_tree.setColumnWidth(1, 200)

        # Create labels for the tree widgets
        converted_label = QLabel("Converted file list")
        corrupt_label = QLabel("Corrupt file list")

        # Create QHBoxLayout for trees and labels
        trees_layout = QHBoxLayout()
        trees_layout.addWidget(self.log_tree)
        trees_layout.addWidget(self.corrupt_tree)

        labels_layout = QHBoxLayout()
        labels_layout.addWidget(converted_label)
        labels_layout.addWidget(corrupt_label)

        # Use QVBoxLayout to arrange the widgets vertically
        layout = QVBoxLayout()
        layout.addWidget(button_group_1)  # Add the first button group
        layout.addLayout(labels_layout)  # Add the labels layout
        layout.addLayout(trees_layout)  # Add the trees layout
        layout.addWidget(button_group_2)  # Add the second button group
        layout.addWidget(self.result_label)
        layout.addWidget(self.progress_bar)

        container = QWidget()
        container.setLayout(layout)
        self.setCentralWidget(container)

        
    def reset(self):
        try:
            self.result_label.clear()
            self.log_tree.clear()
            self.corrupt_tree.clear()
            self.progress_bar.setValue(0)
            self.converted_tree.clear()
        except Exception as e:
            print("Error while clearing converted tree:", str(e))
            
    def scan_f_drive(self):
        root_folder = "F:\\"
        target_folder_name = "Failed"
        error_folders = self.find_error_folders(root_folder, target_folder_name)

        if not error_folders:
            result_text = "No 'ErrorFiles' folders found on the F drive."
        else:
            total_files_found = sum(len(os.listdir(error_folder)) for error_folder in error_folders)
            result_text = f"Total files found on F drive: {total_files_found}"

        self.result_label.setText(result_text)

    def find_error_folders(self, root_folder, target_folder_name):
        error_folders = []
        for root, dirs, files in os.walk(root_folder):
            if target_folder_name in dirs:
                error_folder = os.path.join(root, target_folder_name)
                error_folders.append(error_folder)
        return error_folders
    def time_stamp(self):
        csv_file_path = os.path.join(self.downloads, "error_summary.csv")
        with open(csv_file_path, 'a', newline='') as csv_file:
            csv_writer = csv.writer(csv_file)
            csv_writer.writerow([])
            csv_writer.writerow([])
            current_datetime = datetime.datetime.now().strftime('%A, %Y-%m-%d %H:%M:%S')
            csv_writer.writerow(["Date and Time Stamp:", "", "", current_datetime])
            csv_writer.writerow([])
            
    def create_error_summary_csv(self, error_folders):
        csv_file_path = os.path.join(self.downloads, "error_summary.csv")
        with open(csv_file_path, 'a', newline='') as csv_file:
            csv_writer = csv.writer(csv_file)

            for error_folder in error_folders:
                error_files = [f for f in os.listdir(error_folder) if f.lower().endswith(".pdf")]

                if error_files:
                    csv_writer.writerow(["Folder Location:", "", "", error_folder])
                    csv_writer.writerow(["Files Found:", "", "", len(error_files)])
                    csv_writer.writerow(["File Names:"])
                    for file_name in error_files:
                        csv_writer.writerow(["", "", file_name])

        print(f"Error summary saved to: {csv_file_path}")

    def get_downloads_path(self):
        downloads = os.path.join(os.path.expanduser("~"), "Downloads")
        logs_folder = os.path.join(downloads, 'P2P logs')

        if not os.path.exists(logs_folder):
            os.makedirs(logs_folder)

        return logs_folder

    
    def update_progress(self):
        if self.total_files_to_process == 0:
            progress_percentage = 0
        else:
            progress_percentage = int((self.files_processed / self.total_files_to_process) * 100)
        self.progress_bar.setValue(progress_percentage)

    def process_folders(self, error_folders):
        result_text = ""
        for error_folder in error_folders:
            self.total_files_to_process += len(os.listdir(error_folder))
            self.process_folder(error_folder)
            result_text += f"Processed folder: {error_folder}\n"
        return result_text
        
    def get_downloads_path(self):
        downloads = os.path.join(os.path.expanduser("~"), "Downloads")
        return downloads

    def check_pdf_corruption(self, file_path):
        try:
            with fitz.open(file_path) as doc:
                _ = doc.metadata
            return False
        except Exception as e:
            print(f"Skipping the corrupted file: {file_path}. Error: {str(e)}")
            return True

    def is_pdf_locked(self, pdf_file_path):
        try:
            pdf_document = fitz.open(pdf_file_path)
            return pdf_document.isEncrypted
        except Exception as e:
            print(f"Error checking PDF lock status: {str(e)}")
            return False

    def has_rotation(self, input_path):
        try:
            with open(input_path, 'rb') as pdfFileObj:
                pdfReader = PyPDF2.PdfReader(pdfFileObj)
                for page_num in range(len(pdfReader.pages)):
                    page = pdfReader.pages[page_num]
                    try:
                        if page['/Rotate'] in [90, 270]:
                            return True
                    except KeyError:
                        pass
            return False
        except FileNotFoundError:
            print("File not found.")
            return False


    def construct_output_path(self, base_path, input_path):
        new_file_name = os.path.basename(input_path)[:-4] + '_p2p.pdf'
        parent_directory = os.path.dirname(os.path.dirname(input_path))
        output_path = os.path.join(parent_directory, new_file_name)
        return output_path

    def resize_pages(self, input_path, output_path):
        try:
            parent_directory = os.path.dirname(os.path.dirname(input_path))  
            src = PdfReader(input_path)
            doc = PdfWriter()

            for page_number in range(len(src.pages)):
                pdf_page = doc.add_page(src.pages[page_number])

            output_path = self.construct_output_path(parent_directory, input_path)

            with open(output_path, 'wb') as output_file:
                doc.write(output_file)

            return True
        except Exception as e:
            print(f"Error occurred while reformatting file '{input_path}' using PyPDF2: {str(e)}")
            return False

    def reformat_to_a4_fitz(self, input_path, output_path):
        try:
            parent_directory = os.path.dirname(os.path.dirname(input_path))
            src = fitz.open(input_path)
            doc = fitz.open()

            for ipage in src:
                pdf_page = doc.new_page(width=595, height=842)
                pdf_page.show_pdf_page(pdf_page.rect, src, ipage.number)

            doc.save(output_path)
            doc.close()

        except ValueError as ve:
            if "nothing to show - source page empty" in str(ve):
                print(f"Empty source page error encountered for {input_path}")
                print("Calling resize_pages to reformat the PDF...")
                output_path = self.construct_output_path(parent_directory, input_path)
                self.resize_pages(input_path, output_path)
            else:
                print(f"Error processing {input_path}: {ve}")

    def convert_and_print_to_pdf(self, input_path, output_path):
        try:
            images = convert_from_path(input_path, dpi=200)
            letter_page_size = letter
            c = canvas.Canvas(output_path, pagesize=letter_page_size)
            temp_images = []

            for image in images:
                image_path = f"{output_path.replace('.pdf', '_' + str(images.index(image)) + '.jpg')}"
                image.save(image_path, 'JPEG')
                temp_images.append(image_path)

                c.setPageSize(letter_page_size)
                c.setFont("Helvetica", 10)
                c.drawImage(image_path, 0, 0, width=letter_page_size[0], height=letter_page_size[1])
                c.showPage()

            c.save()

            p = Popen(['print', output_path, 'Microsoft Print to PDF'], stdout=DEVNULL, stderr=DEVNULL, shell=True)
            p.wait()

            for temp_image in temp_images:
                os.remove(temp_image)
                
        except Exception as e:
            print(f"Error occurred while converting and printing to PDF: {str(e)}")


    def process_rotated_pages(self,input_path, output_path, rotation=270):
        if self.check_pdf_corruption(input_path):
            return
        try:
            print("heyyyyyyyyyyyyyyyyyyyyyyyy")
            pdfWriter = PyPDF2.PdfWriter()

            with open(input_path, 'rb') as pdfFileObj:
                pdfReader = PyPDF2.PdfReader(pdfFileObj)
                
                for page_num in range(len(pdfReader.pages)):
                    page = pdfReader.pages[page_num]

                    try:
                        if page['/Rotate'] in [90, 270]:
                            page.rotate(rotation)
                    except KeyError:
                        pass

                    pdfWriter.add_page(page)

            with open(output_path, 'wb') as newFile:
                pdfWriter.write(newFile)

        except Exception as e:
            print(f"Error occurred while processing rotated pages: {str(e)}")
            output_path = self.construct_output_path(parent_directory, input_path)
            self.convert_and_print_to_pdf(input_path, output_path)
            print(f"Original input file processed and saved at: {output_path}")

    def process_folder(self, error_folder):
        error_files = [f for f in os.listdir(error_folder) if f.lower().endswith(".pdf")]
        total_files = len(error_files)
        converted_files_list = []
        corrupt_files_list = []
        processed_files = set()

        for file_name in error_files:
            input_file_path = os.path.join(error_folder, file_name)
            if input_file_path in processed_files:
                continue

            new_file_name = (file_name)[:-4] + "_p2p.pdf"
            output_file_path = self.construct_output_path(error_folder, input_file_path)

            is_corrupt = self.check_pdf_corruption(input_file_path)
            is_locked = self.is_pdf_locked(input_file_path)
            status = None

            if is_corrupt:
                corrupt_files_list.append((file_name, input_file_path, "Corrupted"))
                status = "Corrupted"
                item = QTreeWidgetItem([file_name, status, input_file_path])
                self.corrupt_tree.addTopLevelItem(item)
                self.files_processed += 1
                self.update_progress()
                continue
            if is_locked:
                print(f"Skipping processing of '{input_file_path}' due to encryption.")
                corrupt_files_list.append((error_folder, input_file_path, "Encrypted"))
                self.files_processed += 1
                self.update_progress()
                status = "Encrypted"

                item = QTreeWidgetItem([file_name, status, input_file_path])
                self.corrupt_tree.addTopLevelItem(item)
                continue
            if self.has_rotation(input_file_path):
                try:
                    # Create a temporary file to store the rotated pages
                    with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as temp_file:
                        temp_file_name = temp_file.name
                        temp_file.close()
                    self.process_rotated_pages(input_file_path, temp_file_name)
                    print(f"Rotated pages processed and saved in temporary file: {temp_file_name}")
                    self.reformat_to_a4_fitz(temp_file_name, output_file_path)
                    converted_files_list.append((output_file_path, "Converted"))
                    processed_files.add(input_file_path)
                except Exception as e:
                    print(f"Error occurred while rotating: {str(e)}")
                    status = "Corrupted"
            else:
                print("No rotated pages found in the PDF.")
                try:
                    self.reformat_to_a4_fitz(input_file_path, output_file_path)
                    print(f"File reformatted and saved at: {output_file_path}")
                    converted_files_list.append((output_file_path, "Converted"))
                    processed_files.add(input_file_path)
                except ValueError as ve:
                    print(f"ValueError: {ve}")
                    if self.convert_and_print_to_pdf(input_file_path, output_file_path):
                        print(f"Alternative reformatted and saved at: {output_file_path}")
                        converted_files_list.append((output_file_path, "Converted (Alternative)"))
                        processed_files.add(input_file_path)
                    else:
                        print(f"Alternative reformatting failed for file: {input_file_path}")
                        corrupt_files_list.append((error_folder, input_file_path, "Corrupted"))
                except Exception as e:
                    print(f"Error occurred: {str(e)}")
                    corrupt_files_list.append((error_folder, input_file_path, "Corrupted"))
                    
                










                
##           else:
##              try:
##                  self.reformat_to_a4_fitz(input_file_path, output_file_path)
##                    status = "Converted"
##                except ValueError as ve:
##                  print(f"ValueError: {ve}")
##                    if self.convert_and_print_to_pdf(input_file_path, output_file_path):
##                        print(f"Alternative reformatted and saved at: {output_file_path}")
##                        status = "Converted (Alternative)"
##                    else:
##                        print(f"Alternative reformatting failed for file: {input_file_path}")
##                        status = "Corrupted"
##                except Exception as e:
##                    print(f"Error occurred: {str(e)}")
##                    status = "Corrupted"




        print(f"Total files: {total_files}")
        print(f"Converted files: {len(converted_files_list)}")
        print(f"Corrupted files: {len(corrupt_files_list)}")

        current_datetime = datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')

        converted_files_csv = os.path.join(self.downloads, "converted_files_logs.csv")
        corrupt_files_csv = os.path.join(self.downloads, "Failed_files_logs.csv")

        with open(converted_files_csv, 'a') as f:
            if converted_files_list:
                f.write(f" \n {error_folder},{current_datetime},\n")
                f.write("File Name,Converted Location,Status\n")
            for converted_file in converted_files_list:
                file_name = os.path.basename(converted_file[0])
                f.write(f"{file_name},{converted_file[0]},{converted_file[1]}\n")

        with open(corrupt_files_csv, 'a') as f:
            if corrupt_files_list:
                f.write(f" \n {error_folder},{current_datetime},\n")
                f.write("Folder Location,File Name,Status\n")
                for corrupt_file_info in corrupt_files_list:
                    folder_name = corrupt_file_info[0]
                    file_name = os.path.basename(corrupt_file_info[1])
                    status = corrupt_file_info[2]
                    f.write(f"{folder_name},{file_name},{status}\n")

        print(f"List of converted files saved to: {converted_files_csv}")
        print(f"List of corrupted files saved to: {corrupt_files_csv}")
        self.create_error_summary_csv([error_folder])


    def start_converting(self):
        self.log_tree.clear()

        root_folder = "F:\\"
        target_folder_name = "Failed"
        error_folders = self.find_error_folders(root_folder, target_folder_name)

        if not error_folders:
            result_text = "No 'ErrorFiles' folders found on the F drive."
        else:
            try:
                result_text = self.process_folders(error_folders)
            except Exception as e:
                result_text = f"An error occurred: {str(e)}"
                QMessageBox.critical(self, "Error", f"An error occurred: {str(e)}")

        self.result_label.setText(result_text)


        


    def load_and_display_csv(self, csv_file, status):
        self.log_tree.clear()  # Clear the QTreeWidget

        downloads = self.get_downloads_path()
        csv_path = os.path.join(downloads, csv_file)
        if os.path.exists(csv_path):
            with open(csv_path, 'r') as f:
                lines = f.readlines()[2:]  # Skip the first two lines
                for line in lines:
                    parts = line.strip().split(',')
                    if len(parts) == 3:
                        file_name, original_file_location, _ = parts  # Use original_file_location instead of file_location
                        item = QTreeWidgetItem([file_name, status, original_file_location])  # Pass original_file_location
                        self.log_tree.addTopLevelItem(item)  # Add the item to the tree




if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = PDFConverterApp()
    window.show()
    sys.exit(app.exec_())
