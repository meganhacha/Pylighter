from docx import Document
from docx.shared import Pt
from docx.enum.text import WD_COLOR_INDEX
from tkinter import Tk, filedialog
import os

days_of_the_week = {"MONDAY", "TUESDAY", "WEDNESDAY", "THURSDAY", "FRIDAY", "SATURDAY", "SUNDAY"}
highlight_list = []
bold_list = []
underline_list = []

#Bold event titles and subsequent info if on the same line.
def bold_text(doc, keyword):
   for p in doc.paragraphs:

      for r in p.runs:
         if keyword in r.text:
            r.font.bold = True

#Underline dates
def underline_text(doc, keyword):
   for p in doc.paragraphs:
      
      for r in p.runs:
         if any(word in r.text for word in keyword):
            r.font.underline = True

#Change the font and text size of the entire document.
def allover_text(doc, keyword):
   for p in doc.paragraphs:

      for r in p.runs:
         r.font.size = Pt(12)
         r.font.name = 'Arial'

#Get each individual line, ignoring whitespaces.
def get_paragraph_text(paragraph):
   text = ''.join(run.text for run in paragraph.runs)
   text.strip()
   return text

#Highlight the lines that are applicable, based on the list of keywords below.
def highlight_lines(doc, keyword, color):

   for paragraph in doc.paragraphs:
      if keyword in get_paragraph_text(paragraph):

         for run_h in paragraph.runs:
            run_h.font.highlight_color = color

#Allows the saving of a new -- now highlighted and edited -- Word doc.
def save_doc(doc, output_path):
   doc.save(output_path)

#Asks the user to select the .docx file they would like to perform the highlighting and other edits on.
def select_file():
   root = Tk()
   root.withdraw()

   file_path = filedialog.askopenfilename(
      title = "Select docx File",
      filetypes=[("Word Documents", "*.docx"), ("All Files", "*.*")]
   )

   return file_path

#Using the provided keywords, selects the highlight color for each section.
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
      #Open the selected doc, and add the applicable keywords that need to have their sections highlighted.
      doc= Document(input_file)
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

      #Traverse the document, highlighting each word with its corresponding color.
      for key in highlight_list:
         color = get_highlight(key)
         highlight_lines(doc, key, color)

      #Create a list of words that are typically bolded in the document.
      bold_list.append('Organization:')
      bold_list.append('Contact person day of event:')
      bold_list.append('# of people expected to attend:')
      bold_list.append('Event Description:')
      bold_list.append('Setup/Tear Down:')
      bold_list.append('Tech Needs:')
      bold_list.append('Food Services:')
      bold_list.append('Security')

      #Change the text to the correct font and text size.
      allover_text(doc, bold_list)

      #Use the bold list to bold all instances of the bold list in the Document.
      for b_word in bold_list:
         bold_text(doc, b_word)
      for h_word in highlight_list:
         bold_text(doc, h_word)

        
      

      output_file = filedialog.asksaveasfilename(
         defaultextension=".docx",
         filetypes=[("Word Documents", "*.docx"), ("All Files", "*.*")]

       )

      if output_file:
         save_doc(doc, output_file)
         print(f"Document saved to {output_file}")
      else:
         print("No output file selected. Exiting.")