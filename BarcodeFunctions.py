import os
import pandas as pd
import numpy as np
import reportlab
import barcode
from pathlib import Path
from barcode import writer
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter
from reportlab.graphics import renderPDF
from svglib.svglib import svg2rlg
from IPython.display import SVG, display
from reportlab.pdfbase.pdfmetrics import stringWidth, getAscentDescent


# Turns serial number into svg barcode

def serial_to_barcode(serial):
  if serial == "-":
    return False, None
  code128 = barcode.get('code128', serial)
  filename = code128.save('code128', options={"font_size": 0, "quiet_zone": 0})
  return True, filename

# Combines two lists so each same index is now a smaller list unit

def combined_list(list1, list2):
  final_tuple_list = []
  if len(list1) != len(list2):
    raise IndexError("List of serial numbers and titles do not line up")
  for i in range(len(list1)):
    final_tuple_list.append((list1[i], list2[i]))
  return final_tuple_list

# Add "-" values until the list length is divisible by 21

def divisible_by_21(label_list):
    if len(label_list) % 21 != 0:
      label_list.append("-")
      divisible_by_21(label_list)
    return label_list

# Turns CSV into a numpy array
# The first index in each smallest tuple will be the serial number and second will be the title (ex. Cello 1/4)

def csv_to_padded_array(csv_path):
  temp = pd.read_csv(csv_path)
  column_serial = temp['Serial Number'].tolist()
  column_title = temp['Title'].tolist()
  if len(column_serial) != len(column_title):
    raise IndexError("List of serial numbers and titles do not line up")
  column_serial = divisible_by_21(column_serial)
  column_title = divisible_by_21(column_title)
  final_columns = combined_list(column_serial, column_title)
  labels = np.array(final_columns)
  reshaped_labels = np.reshape(labels, (-1, 7, 3, 2))
  return reshaped_labels

# Gets matrix indices for everything

def barcode_coords(x, y):
  y_axis = 18 + 18 + 108 * (6) - 108 * (y)
  x_axis = 30 + 204 * (x)
  return x_axis, y_axis

def serial_coords(x, y):
  y_axis = 18 + 9 + 108 * (6) - 108 * (y)
  x_axis = 102 + 204 * (x)
  return x_axis, y_axis

def title_coords(x, y):
  # y_axis = 18 + (108 - 18) + 108 * (y) # Also Equal to 108 * (y + 1) but written in this form for reading
  y_axis = 18 + 108 - 9 + 108 * (6) - 108 * (y)
  x_axis = 102 + 204* (x)
  return x_axis, y_axis

# Draws image of barcode onto PDF

def draw_on_pdf(barcode_path, x, y, doc):
  drawing = svg2rlg(barcode_path)
  target_width = 2*72
  target_height = 1*72
  scale_x = target_width / drawing.width
  scale_y = target_height / drawing.height
  drawing.scale(scale_x, scale_y)
  renderPDF.draw(drawing, doc, x, y)

# Writes strings onto PDF

def write_on_pdf(text, x, y, doc, font_name="Helvetica", font_size="30"):
  doc.setFont(font_name, font_size)
  ascent, descent = getAscentDescent(font_name)
  text_height = (ascent - descent)/1000 * font_size
  adjusted_y = y - text_height/2
  doc.drawCentredString(x, adjusted_y, text)

# Main Pipline (From CSV to PDF):

def main(csv_path, pdf_path=""):
  if pdf_path == "":
    current_folder = os.path.dirname(os.path.abspath(__file__))
    pdf_path = current_folder + "/barcodes.pdf"
  print("Reading CSV...")
  serial_nums = csv_to_padded_array(csv_path)
  print("CSV loaded. Creating canvas...")
  doc = canvas.Canvas(pdf_path, pagesize=letter)
  for z in range(len(serial_nums)):
    print(f"Processing page {z + 1}/{len(serial_nums)}")
    for y in range(len(serial_nums[0])):
      for x in range(len(serial_nums[0][0])):
        serial = str(serial_nums[z][y][x][0])
        boolean, barcode_path = serial_to_barcode(serial)
        if boolean:
          title = str(serial_nums[z][y][x][1])
          bar_x, bar_y= barcode_coords(x, y)
          title_x, title_y = title_coords(x, y)
          serial_x, serial_y = serial_coords(x, y)
          draw_on_pdf(barcode_path, bar_x, bar_y, doc)
          write_on_pdf(title, title_x, title_y, doc, font_size=10)
          write_on_pdf(serial, serial_x, serial_y, doc, font_size=10)
    doc.showPage()
  print("Saving PDF...")
  doc.save()
  print("Done!")
