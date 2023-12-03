from docx import Document
from docx.enum.text import WD_COLOR_INDEX
from tkinter import Tk, filedialog
import os

days_of_the_week = {"MONDAY", "TUESDAY", "WEDNESDAY", "THURSDAY", "FRIDAY", "SATURDAY", "SUNDAY"}

def highlight_lines(doc, keyword, color):

    for paragraph in doc.paragraphs:
        if paragraph.text.startswith(keyword):
           for run in paragraph.runs:
             run.font.highlight_color = color

def save_doc(doc, output_path):
   doc.save(output_path)

def select_file():
   root = Tk()
   root.withdraw()

   file_path = filedialog.askopenfilename(
      title = "Select docx File",
      filetypes=[("Word Documents", "*.docx"), ("All Files", "*.*")]
   )

   return file_path

def get_highlight(keyword):
    key = keyword
    if key == "Event:" or key == "Meeting:":
      highlight_color = WD_COLOR_INDEX.RED
      return highlight_color
    
    elif key == "Event Date/Time:" or key == "Date/Time:":
       highlight_color = WD_COLOR_INDEX.YELLOW
       return highlight_color
    
    elif key == "Location:":
       highlight_color = WD_COLOR_INDEX.TURQUOISE
       return highlight_color

    elif key in days_of_the_week:
       highlight_color = WD_COLOR_INDEX.PINK
       return highlight_color
    
    elif key == "Concierge":
       highlight_color = WD_COLOR_INDEX.BRIGHT_GREEN
       return highlight_color


if __name__ == "__main__":
   
    input_file = select_file()

    if not input_file:
      print("No file selected. Exiting")
    else:

        doc= Document(input_file)
        highlight_list = []
        highlight_list.append( "Event:")
        highlight_list.append("Meeting:")
        highlight_list.append("Event Date/Time:")
        highlight_list.append("Date/Time:")
        highlight_list.append("Location:")
        highlight_list.append("MONDAY")
        highlight_list.append("TUESDAY") 
        highlight_list.append("WEDNESDAY")
        highlight_list.append("THURSDAY")
        highlight_list.append("FRIDAY")
        highlight_list.append("SATURDAY")
        highlight_list.append("SUNDAY")
        highlight_list.append("Concierge")

        for key in highlight_list:
           color = get_highlight(key)
           print(f"{color}")
           highlight_lines(doc, key, color)

        output_file = filedialog.asksaveasfilename(
          defaultextension=".docx",
          filetypes=[("Word Documents", "*.docx"), ("All Files", "*.*")]

       )

        if output_file:
          save_doc(doc, output_file)
          print(f"Document saved to {output_file}")
        else:
          print("No output file selected. Exiting.")