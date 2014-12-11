import sys
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
import pylab as pl

kpi = 9
font_size = 7
bins = {}

def build_bins():
  global bins
  bins[0] = { 'limits': [0,2], 'label': '0-2', 'count': 0 }
  bins[1] = { 'limits': [2,4], 'label': '2-4', 'count': 0 }
  bins[2] = { 'limits': [4,8], 'label': '4-8', 'count': 0 }
  bins[3] = { 'limits': [8,16], 'label': '8-16', 'count': 2 }
  bins[4] = { 'limits': [16,32], 'label': '16-32', 'count': 0 }

# This function adds the label to the right of the bars
def autolabel(rects):
    distance_to_rect = 0.25
    for i, rect in enumerate(rects):
        width = rect.get_width()
        label_text = str(round(float(width), 2)) 
        label_text = width 
        plt.text(width + distance_to_rect, rect.get_y() + rect.get_height() / 2., label_text, ha="left", va="center", fontsize=font_size)

def create_plot(image_format, dpi, division=None):
  global bins
  labels = [bins[i]['label'] for i in bins.keys()]
  values = [bins[i]['count'] for i in bins.keys()]
  fig = plt.figure(figsize=(2,1))

  # Plot the red horizontal bars and set labels
  rects = plt.barh(range(len(values)), values, height=0.7, align="center", color="#8A0707", edgecolor="none")
  autolabel(rects)

  # x-tick and y-tick labels (no x-tick labels)
  pl.yticks(range(len(labels)), labels, fontsize=font_size)
  pl.xticks(range(0, 5, 1), [""])

  # Set axis labels
  pl.ylabel('Scaling Factor', fontsize=font_size)
  pl.xlabel('No. projects scaled in CPUs (total=%s)' % sum(values), fontsize=font_size)

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
  name = 'pics/KPI_sample.%s' % image_format
  pl.savefig(name, bbox_inches='tight', dpi=dpi)
  plt.close()
  print 'Saved %s' % name


def create_sample_plot(image_format, dpi):
  global bins
  build_bins()
  create_plot(image_format, dpi)
