import sys
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
import pylab as pl

kpi = 9
font_size = 7
scaling_types = { 1: 'Throughput', 2: 'CPUs', 3: 'Memory', 4: 'Optimisation' }
bins = {}

def build_bins():
  ''' We report scaling factors in bins. '''
  global bins
  for i in range(0,11):
    lower_limit = 2**i
    upper_limit = 2**(i+1)
    if lower_limit == 1:
      lower_limit = 0
    bins[i] = { 'limits': [lower_limit, upper_limit], 'label': '%s-%s' % (lower_limit, upper_limit), 'count': 0 }

def get_bin_index(scaling_factor):
  ''' Get the index of a bin for a given scaling factor '''
  global bins
  for key in bins.keys():
    if scaling_factor >= bins[key]['limits'][0] and scaling_factor < bins[key]['limits'][1]:
      return key
  raise Exception('No bin found for value: %s' % scaling_factor)

def nullify_bins():
  global bins
  for key in bins:
    bins[key]['count'] = 0

def query(db, query_string):
  db.query(query_string)
  r=db.store_result()
  result=r.fetch_row(maxrows=0,how=1)
  return result

def populate_bins_with_kpis(db, kpi, type, institution, oldest_date, division=None):
  global bins
  q = "SELECT DISTINCT projectId from project_kpi pk "
  q += "INNER JOIN project p ON p.id = pk.projectId "
  q += "WHERE pk.date >= '%s' " % oldest_date
  q += "AND p.hostInstitution = '%s' " % institution
  if division:
    q += "AND p.division = '%s'" % division
  results = query(db, q)
  if results:
    pids = [r['projectId'] for r in results]
    for pid in pids:
      q = "SELECT MAX(value) AS maxval FROM project_kpi "
      q += "WHERE projectId = %s " % pid
      q += "AND date >= '%s' " % oldest_date
      q += "AND code = %s" % type
      result = query(db, q)
      if result:
        if len(result) == 1:
          maxval = result[0]['maxval']
          if maxval and float(maxval) > 0:
            bins[get_bin_index(float(maxval))]['count'] += 1
        else:
          raise Exception('Expecting only one result here')

def get_divisions(db, institution_name):
  q = '''SELECT name FROM division WHERE institutionId = (SELECT id FROM institution WHERE name = '%s')''' % institution_name
  results = query(db, q)
  return [r['name'] for r in results]

# This function adds the label to the right of the bars
def autolabel(rects):
    max_width = max([r.get_width() for r in rects])
    distance_to_rect = 0.25 * (1 + max_width/10)
    for i, rect in enumerate(rects):
        width = rect.get_width()
        label_text = str(round(float(width), 2)) 
        label_text = width 
        plt.text(width + distance_to_rect, rect.get_y() + rect.get_height() / 2., label_text, ha="left", va="center", fontsize=font_size)

def create_plots(db, institution, oldest_date, image_format, dpi, division=None):
  global bins
  for type in scaling_types.keys():
    nullify_bins()
    populate_bins_with_kpis(db, kpi, type, institution, oldest_date, division)
    labels = [bins[i]['label'] for i in bins.keys()]
    values = [bins[i]['count'] for i in bins.keys()]
    fig = plt.figure(figsize=(2,2.3))

    # Plot the red horizontal bars and set labels
    rects = plt.barh(range(len(values)), values, height=0.7, align="center", color="#8A0707", edgecolor="none")
    autolabel(rects)

    # x-tick and y-tick labels (no x-tick labels)
    pl.yticks(range(len(labels)), labels, fontsize=font_size)
    pl.xticks(range(0, 5, 1), [""])

    # Set axis labels
    pl.ylabel('Scaling Factor', fontsize=font_size)
    pl.xlabel('No. projects scaled in %s (total=%s)' % (scaling_types[type],sum(values)), fontsize=font_size)

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
    if division:
      name = 'pics/KPI_%s_%s_%s' % (str(institution).replace(' ','-'), str(division).replace(' ','-'), scaling_types[type])
    else:
      name = 'pics/KPI_%s_%s' % (str(institution).replace(' ','-'), scaling_types[type])
    name += '.%s' % image_format
    pl.savefig(name, bbox_inches='tight', dpi=dpi)
    plt.close()
    print 'Saved %s' % name


def create_kpi_plots(db, institution, oldest_date, image_format, dpi):
  global bins
  build_bins()
  create_plots(db, institution, oldest_date, image_format, dpi)
  for division in get_divisions(db, institution):
    create_plots(db, institution, oldest_date, image_format, dpi, division)
