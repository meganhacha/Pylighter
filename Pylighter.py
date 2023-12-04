from docx import Document
from docx.shared import Pt
from docx.enum.text import WD_COLOR_INDEX
from tkinter import Tk, filedialog
import os

days_of_the_week = {"MONDAY", "TUESDAY", "WEDNESDAY", "THURSDAY", "FRIDAY", "SATURDAY", "SUNDAY"}

def bold_text(doc, keyword):
   for p in doc.paragraphs:

      for r in p.runs:
         if any(word in r.text for word in keyword):
            r.font.bold = True

def underline_text(doc, keyword):
   for p in doc.paragraphs:
      
      for r in p.runs:
         if any(word in r.text for word in keyword):
            r.font.underline = True

def allover_text(doc):
   for p in doc.paragraphs:

      for r in p.runs:
         r.font.size = Pt(12)
         r.font.name = 'Arial'

def get_paragraph_text(paragraph):
   text = ''.join(run.text for run in paragraph.runs)
   text.strip()
   return text

def highlight_lines(doc, keyword, color):

   for paragraph in doc.paragraphs:
      if keyword in get_paragraph_text(paragraph):

         for run_h in paragraph.runs:
            run_h.font.highlight_color = color


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
   highlight_color = WD_COLOR_INDEX.GRAY_50

   if key in days_of_the_week:
      highlight_color = WD_COLOR_INDEX.PINK
      return highlight_color
   
   if key == "Event:" or key == "Meeting:" or key == "SETUP":
      highlight_color = WD_COLOR_INDEX.RED
      return highlight_color
   
   if key == "Location:":
      highlight_color = WD_COLOR_INDEX.TURQUOISE
      return highlight_color
   
   if key == "Event Date/Time:" or key == "Date/Time:":
      highlight_color = WD_COLOR_INDEX.YELLOW
      return highlight_color
   
   if key == "Concierge":
      highlight_color = WD_COLOR_INDEX.BRIGHT_GREEN
      return highlight_color
   
   return highlight_color


if __name__ == "__main__":
   
    input_file = select_file()

    if not input_file:
      print("No file selected. Exiting")
    else:

      doc= Document(input_file)
      highlight_list = []
      highlight_list.append("Event:")
      highlight_list.append("Meeting:")
      highlight_list.append("SETUP")
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
         highlight_lines(doc, key, color)
      
      bold_list = []
      bold_list.append('Organization:')
      bold_list.append('Contact person day of event:')
      bold_list.append('# of people expected to attend:')
      bold_list.append('Event Description:')
      bold_list.append('Setup/Tear Down:')
      bold_list.append('Tech Needs:')
      bold_list.append('Food Services:')
      bold_list.append('Security')

      allover_text(doc)
      print(highlight_list[0])
      bold_text(doc, highlight_list)
      bold_text(doc, bold_list)

        
      

      output_file = filedialog.asksaveasfilename(
         defaultextension=".docx",
         filetypes=[("Word Documents", "*.docx"), ("All Files", "*.*")]

       )

      if output_file:
         save_doc(doc, output_file)
         print(f"Document saved to {output_file}")
      else:
         print("No output file selected. Exiting.")