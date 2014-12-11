import os
import _mysql
import docx
from docx.oxml.shared import OxmlElement, qn
from docx_research_output import add_research_output
from docx_feedback import add_feedback
from docx_kpis import add_kpis
from plot_research_output import create_research_output_plots
from plot_kpi import create_kpi_plots
from plot_kpi_sample import create_sample_plot
import db

doc_name = 'CeR_Cluster_Report_2014.docx'
doc_title = 'Centre for eResearch Cluster Report'
institution = 'University of Auckland'
oldest_date = '2013-07-05'

# image parameters
image_format = 'png'
dpi = 360

# create db connection
db = _mysql.connect(**db.config)

# create document
doc = docx.Document()

try:
  if not os.path.exists('pics'):
    os.makedirs('pics')
  create_research_output_plots(db, institution, oldest_date, image_format, dpi)
  create_kpi_plots(db, institution, oldest_date, image_format, dpi)
  create_sample_plot(image_format, dpi)
  doc.add_heading(doc_title, 0)
  doc.add_page_break()
  add_research_output(db, doc, institution, oldest_date)
  doc.add_page_break()
  add_kpis(db, doc, institution, oldest_date)
  doc.add_page_break()
  add_feedback(db, doc, institution, oldest_date)
  doc.save(doc_name)
finally:
  db.close()
  
