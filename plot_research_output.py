import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
import pylab as pl

font_size=8
research_output = {}

def get_divisions(db, institution_name):
  q = '''SELECT name FROM division WHERE institutionId = (SELECT id FROM institution WHERE name = '%s')''' % institution_name
  results = query(db, q)
  return [r['name'] for r in results]

def query(db, query_string):
  db.query(query_string)
  r=db.store_result()
  result=r.fetch_row(maxrows=0,how=1)
  return result

def init_research_output_types(db):
  global research_output
  q = 'SELECT id,name from researchoutputtype'
  res = query(db, q)
  for r in res:
    research_output[int(r['id'])] = { 'name': r['name'], 'count': 0 }
  return research_output 

def nullify_research_outputs():
  global research_output
  for k in research_output.keys():
    research_output[k]['count'] = 0

def get_research_outputs(db, oldest_date, institution, division=None):
  q = 'SELECT DISTINCT ro.id, ro.typeId FROM researchoutput ro \
       INNER JOIN researcher_project rp ON rp.projectId=ro.projectId \
       INNER JOIN researcher r ON r.id=rp.researcherId '
  if institution:
    q += 'AND r.institution=\'%s\' ' % institution
    if division:
      q += 'AND r.division=\'%s\' ' % division
  q += 'WHERE date > \'%s\' ' % oldest_date
  q += ' AND  rp.researcherRoleId=1'
  results = query(db, q)
  return [int(r['typeId']) for r in results]

def get_bin_index(value):
  val = round(float(value))
  for key in bins.keys():
    if val >= bins[key]['limits'][0] and val <= bins[key]['limits'][1]:
      return key
  raise Exception('outside of bin: %s' % value)

# This function adds the deaths per minute label to the right of the bars
def autolabel(rects):
    max_width = max([r.get_width() for r in rects])
    distance_to_rect = 0.1 * (1 + max_width/15)
    for i, rect in enumerate(rects):
        width = rect.get_width()
        label_text = str(round(float(width), 2))
        label_text = width
        plt.text(width + distance_to_rect, rect.get_y() + rect.get_height() / 2., label_text, ha="left", va="center", fontsize=font_size)


def create_plot(db, oldest_date, institution, image_format, dpi, division=None):
  nullify_research_outputs()
  ros = get_research_outputs(db, oldest_date, institution, division)
  for r in ros:
    research_output[r]['count'] += 1

  # sort by count
  ro = sorted([research_output[x] for x in research_output.keys()],key=lambda r: r['count'], reverse=False)
  labels = [x['name'] for x in ro]
  counts = [x['count'] for x in ro]

  fig = plt.figure(figsize=(5,3))

  # Plot the red horizontal bars
  rects = plt.barh(range(len(counts)), counts, height=0.7, align="center", color="#084B8A", edgecolor="none")
  autolabel(rects)

  # x tick and y tick labels (no x tick labels)
  pl.yticks(range(len(labels)), labels, fontsize=font_size)
  pl.xticks(range(0, 5, 1), [""])

  # Set axis labels
  pl.ylabel('Research Output Type', fontsize=font_size)
  pl.xlabel('No. Research Outputs (total=%s)' % sum(counts), fontsize=font_size)

  # Plot styling
  # Remove the plot frame lines
  ax = pl.axes()
  ax.spines["top"].set_visible(False)
  ax.spines["right"].set_visible(False)
  ax.spines["left"].set_visible(True)
  ax.spines["bottom"].set_visible(False)

  # Hide ticks
  ax.tick_params(which='both', left='off', bottom='off', right='off', top='off')

  # Save the figure as a PNG
  # Save the figure as a PNG
  if division:
    name = 'pics/ResearchOutputs_%s_%s' % (str(institution).replace(' ','-'), str(division).replace(' ','-'))
  else:
    name = 'pics/ResearchOutputs_%s' % (str(institution).replace(' ','-'))
  name += '.%s' % image_format
  pl.savefig(name, bbox_inches='tight', dpi=dpi)
  plt.close()
  print 'Saved %s' % name

def create_research_output_plots(db, institution, oldest_date, image_format, dpi):
  research_output = init_research_output_types(db)
  create_plot(db, oldest_date, institution, image_format, dpi)
  for division in get_divisions(db, institution):
    create_plot(db, oldest_date, institution, image_format, dpi, division)

