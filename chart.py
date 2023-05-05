#!/usr/bin/env python
# _*_ coding:utf-8 _*_

import warnings
from io import BytesIO
from operator import itemgetter

import matplotlib

matplotlib.use("Agg")
import matplotlib.cm as cm
import matplotlib.patches as patches
import matplotlib.pyplot as plt
import numpy as np
import pandas as pd
from matplotlib import colors as mcolors
from matplotlib import rcParams
from matplotlib.ticker import PercentFormatter, AutoMinorLocator
from openpyxl.drawing.image import Image

from Monthly_parser import chk_month_order, monthly_data_fill
from Weekly_parser import get_week_ooc

warnings.simplefilter(action='ignore', category=UserWarning)
rcParams['font.family'] = 'sans-serif'
rcParams['font.sans-serif'] = ['Tahoma']
c = mcolors.hex2color
bar_width = 0.5


def weird_division(n, d):
    return n / d if d else 0


def segment(lst):
    """evenly"""
    ct = 0
    n = len(lst)
    it = iter(lst)
    yield next(it)
    for x in it:
        ct += 1
        yield ct / n
        yield x


def chk_last_row(df, col, content):
    try:
        # result = bool(df.at[df.last_valid_index(), col] == content)
        result = bool(df[col].iloc[-1] == content)
    # except KeyError:  # even col is not in df, return False
    except Exception:  # even col is not in df, return False
        result = False
    return result


def cm2inch(value):
    return value / 2.54


def split_list_right(a_list):
    half = len(a_list) // 2
    return a_list[half:]


def slicer(start, end, spec=100):
    result = np.linspace(start, end, spec).tolist()
    return result


def gradient(startcolor, endcolor):
    start_color = c(startcolor)
    end_color = c(endcolor)
    r, g, b = start_color
    re, ge, be = end_color
    midcolor = ((r + re) / 2, (g + ge) / 2, (b + be) / 2)
    # grad = [np.array([a]) for a in zip(slicer(r, re), slicer(g, ge), slicer(b, be))]
    grad = [a for a in zip(slicer(r, re), slicer(g, ge), slicer(b, be))]
    downgrad = grad[::2]
    midgrad = list(reversed(downgrad)) + downgrad
    return {'gradient': grad, 'midcolor': midcolor, 'middle grad': midgrad}


default_grad = gradient('#993366', '#47182F')


# def gradientbars(bars, grad=default_grad['gradient']):
def gradientbars(bars, rot90=False):
    ax = bars[0].axes
    lim = ax.get_xlim() + ax.get_ylim()
    grad = np.atleast_2d(np.linspace(0, 1, 256)).T
    if rot90:
        grad = np.rot90(grad, axes=(-2, -1))
    for bar in bars:
        bar.set_zorder(1)
        bar.set_facecolor("none")
        x, y = bar.get_xy()
        w, h = bar.get_width(), bar.get_height()
        ax.imshow(grad, cmap='rot90' if rot90 else 'pareto', extent=[x, x + w, y, y + h], aspect="auto", zorder=0)
    ax.axis(lim)


def make_colormap(seq):
    """Return a LinearSegmentedColormap
    seq: a sequence of floats and RGB-tuples. The floats should be increasing
    and in the interval (0,1).
    """
    seq = [(None,) * 3, 0.0] + list(seq) + [1.0, (None,) * 3]
    cdict = {'red': [], 'green': [], 'blue': []}
    for i, item in enumerate(seq):
        if isinstance(item, float):
            r1, g1, b1 = seq[i - 1]
            r2, g2, b2 = seq[i + 1]
            cdict['red'].append([item, r1, r2])
            cdict['green'].append([item, g1, g2])
            cdict['blue'].append([item, b1, b2])
    return mcolors.LinearSegmentedColormap('CustomMap', cdict)


fsvcolors = [c('#9999FF'), c('#00FFFF'), c('#993366'), c('#008080'), c('#590059'), c('#FF00FF'), c('#FF8080'),
             c('#0066CC'),
             c('#000080'), c('#CCFFCC'), c('#FFFF00'), c('#008000'), c('#800000')]
bsvcolors = [c('#9999FF'), c('#993366'), c('#FFFFCC'), c('#CCFFFF'), c('#FF00FF'), c('#00FFFF'), c('#660066'),
             c('#FF8080'),
             c('#0066CC'), c('#CCCCFF'), c('#FFFF00'), c('#000080')]
cpcolors = [c('#9999FF'), c('#993366'), c('#FFFFCC'), c('#CCFFFF'), c('#CCCCFF'), c('#FF00FF'), c('#660066')]
monthly_colors = [c('#4BACC6'), c('#8064A2'), c('#4F81BD')]
colordict = {'FSV': fsvcolors, 'BSV': bsvcolors, 'CP': cpcolors}
paretocolors = default_grad['gradient']
rot90colors = default_grad['middle grad']
# fsvcmap = make_colormap(list(segment(fsvcolors)))
# cm.register_cmap(name='FSV', cmap=fsvcmap)
# bsvcmap = make_colormap(list(segment(bsvcolors)))
# cm.register_cmap(name='BSV', cmap=bsvcmap)
# cpcmap = make_colormap(list(segment(cpcolors)))
# cm.register_cmap(name='CP', cmap=cpcmap)
paretocmap = make_colormap(list(segment(paretocolors)))
cm.register_cmap(name='pareto', cmap=paretocmap)
rot90cmap = make_colormap(list(segment(rot90colors)))
cm.register_cmap(name='rot90', cmap=rot90cmap)


def two_list_to_df(lst_data, lst_labels):
    data_dict = {label: [data] for label, data in zip(lst_labels, lst_data)}
    result = pd.DataFrame(data_dict)
    return result


def pareto(data, labels=None, cumplot=True, axes=None, limit=1.0, rot90=False,
           data_args=(), data_kw=None, line_args=(), line_kw=None, grad=None,
           limit_kw=None, not_others_num=None, input_percent=True, xticklabel_rot=None):
    """
    Plots a `pareto chart`_ of input categorical data. NOTE: The matplotlib
    command ``show()`` will need to be called separately. The default chart
    uses the following styles:

    - bars:
       - color = blue
       - align = center
       - width = 0.9
    - cumulative line:
       - color = blue
       - linestyle = solid
       - markers = None
    - limit line:
       - color = red
       - linestyle = dashed

    Parameters
    ----------
    data : array-like
        The categorical data to be plotted (not necessary to put in descending
        order prior). or one line Dataframe

    Optional
    --------
    labels : list
        A list of strings of the same length as ``data`` that provide labels
        to the categorical data. If none provided, a simple integer value is
        used to label the data, relating to the original order as given. If
        a list is provided, but not the same length as ``data``, then it will
        be treated as if no labels had been input at all.
    cumplot : bool
        If ``True``, a cumulative percentage line plot is included in the chart
        (Default: True) and a second axis indicating the percentage is returned.
    axes : axis object(s)
        If valid matplotlib axis object(s) are given, the chart and cumulative
        line plot of placed on the given axis. Otherwise, a new figure is
        created.
    limit : scalar
        The cumulative percentage value at which the input data should be
        "chopped off" (should be a value between zero and one).
    data_args : tuple
        Any valid ``matplotlib.pyplot.bar`` non-keyword arguments to apply to
        the bar chart.
    data_kw : dict
        Any valid ``matplotlib.pyplot.bar`` keyword arguments to apply to
        the bar chart.
    line_args : tuple
        Any valid ``matplotlib.pyplot.plot`` non-keyword arguments to apply to
        the cumulative line chart.
    line_kw : dict
        Any valid ``matplotlib.pyplot.plot`` keyword arguments to apply to
        the cumulative line chart.
    limit_kw : dict
        Any valid ``matplotlib.axes.axhline`` keyword arguments to apply to
        the limit line.
    not_others_num : default None, no others group in pareto chart,
        if int, keep top num of data and sum the rest as others group

    input_percent: default True, change input data to percent format

    xticklabel_rot: rotate the x axis lable, default None

    grad: gradientbars function, default None

    Returns
    -------
    fig : matplotlib.figure
        The parent figure object of the chart.
    ax1 : matplotlib.axis
        The axis for the categorical data.
    ax2 : matplotlib.axis
        The axis for the cumulative line plot (not returned if
        ``cumplot=False``).

    Examples
    --------

    The following code is the same test code if the ``paretoplot.py`` file is
    run with the command-line call ``$ python paretoplot.py``::
        # plot data using the indices as labels
        data = [21, 2, 10, 4, 16]

        # define labels
        labels = ['tom', 'betty', 'alyson', 'john', 'bob']

        # create a grid of subplots
        fig,axes = plt.subplots(2, 2)

        # plot first with just data
        pareto(data, axes=axes[0, 0])
        plt.title('Basic chart without labels', fontsize=10)

        # plot data and associate with labels
        pareto(data, labels, axes=axes[0, 1], limit=0.75, line_args=('g',))
        plt.title('Data with labels, green cum. line, limit=0.75', fontsize=10)

        # plot data and labels, but remove lineplot
        pareto(data, labels, cumplot=False, axes=axes[1, 0],
               data_kw={'width': 0.5, 'color': 'g'})
        plt.title('Data without cum. line, bar width=0.5', fontsize=10)

        # plot data cut off at 95%
        pareto(data, labels, limit=0.95, axes=axes[1, 1], limit_kw={'color': 'y'})
        plt.title('Data trimmed at 95%, yellow limit line', fontsize=10)

        # format the figure and show
        fig.canvas.set_window_title('Pareto Plot Test Figure')
        plt.show()
    .. _pareto chart: http://en.wikipedia.org/wiki/Pareto_chart

    """
    if labels is None:
        labels = []
    if data_kw is None:
        data_kw = {}
    if line_kw is None:
        line_kw = {}
    if limit_kw is None:
        limit_kw = {}
    if isinstance(data, pd.DataFrame):
        if len(data) > 1:
            data = data.T
        labels = list(data)
        data = data.iloc[-1].tolist()
        # data = data.tail(1).values.tolist()

    # re-order the data in descending order
    data = list(data)
    n = len(data)
    if n != len(labels):
        labels = range(n)
    ordered = sorted(zip(data, labels), key=itemgetter(0), reverse=True)
    ordered_data = [dat for dat, lab in ordered]
    ordered_labels = [lab for dat, lab in ordered]
    if not_others_num is not None:
        ordered_labels = ordered_labels[:not_others_num]
        ordered_labels.append('Others')
        others = sum(ordered_data[not_others_num:])
        ordered_data = ordered_data[:not_others_num]
        ordered_data.append(others)
        n = len(ordered_data)
    # allow trimming of data (e.g. 'limit=0.95' keeps top 95%)
    assert 0.0 <= limit <= 1.0, 'limit must be a positive scalar between 0.0 and 1.0'

    # create the cumulative line data
    line_data = [0.0] * n
    total_data = float(sum(ordered_data))
    for i, dat in enumerate(ordered_data):
        if i == 0:
            # line_data[i] = dat / total_data
            line_data[i] = weird_division(dat, total_data)
        else:
            # line_data[i] = sum(ordered_data[:i + 1]) / total_data
            line_data[i] = weird_division(sum(ordered_data[:i + 1]), total_data)

    # determine where the data will be trimmed based on the limit
    ltcount = 0
    for ld in line_data:
        if ld < limit:
            ltcount += 1
    limit_loc = range(ltcount + 1)
    try:
        limited_data = [ordered_data[i] for i in limit_loc]
    except IndexError:  # all data is zero
        limit_loc = range(ltcount)
        limited_data = [ordered_data[i] for i in limit_loc]

    limited_labels = [ordered_labels[i] for i in limit_loc]
    limited_line = [line_data[i] for i in limit_loc]
    df_sorted = two_list_to_df(limited_data, limited_labels)
    # if axes is specified, grab it and focus on its parent figure; otherwise,
    # create a new figure
    if axes:
        plt.sca(axes)
        ax1 = axes
        fig = plt.gcf()
    else:
        fig = plt.gcf()
        ax1 = plt.gca()
    fig.set_size_inches(cm2inch(10), cm2inch(11))

    # create the second axis
    if cumplot:
        ax2 = ax1.twinx()

    # plotting
    if 'align' not in data_kw:
        data_kw['align'] = 'center'
    if 'width' not in data_kw:
        data_kw['width'] = 0.9
    bars = ax1.bar(limit_loc, limited_data, *data_args, **data_kw)
    label1 = data_kw.get('label', '')
    ax1.set_ylabel(label1)
    ax1.yaxis.set_tick_params(direction='in')

    if cumplot:
        ax2.plot(limit_loc, [ld * 100 for ld in limited_line], *line_args,
                 **line_kw)
        label2 = line_kw.get('label', '')
        ax2.set_ylabel(label2)
        ax2.yaxis.set_tick_params(direction='in')

    ax1.set_xticks(limit_loc)
    minor_locator = AutoMinorLocator(2)
    ax1.xaxis.set_minor_locator(minor_locator)
    ax1.xaxis.set_tick_params(bottom=False, top=False, which='major')
    ax1.xaxis.set_tick_params(direction='in', which='minor')
    ax1.set_xlim(-0.5, len(limit_loc) - 0.5)

    # formatting
    if cumplot:
        # since the sum-total value is not likely to be one of the tick marks,
        # let's make it the top-most one, regardless of label closeness
        ax1.set_ylim(0, total_data)
        loc = ax1.get_yticks()
        newloc = [loc[i] for i in range(len(loc)) if loc[i] <= total_data]
        newloc += [total_data]
        ax1.set_yticks(newloc)
        ax2.set_ylim(0, 100)
        if limit < 1.0:
            xmin, xmax = ax1.get_xlim()
            if 'linestyle' not in limit_kw:
                limit_kw['linestyle'] = '--'
            if 'color' not in limit_kw:
                limit_kw['color'] = 'r'
            ax2.axhline(limit * 100, xmin - 1, xmax - 1, **limit_kw)

    # set the x-axis labels
    if xticklabel_rot:
        ax1.set_xticklabels(limited_labels, rotation=xticklabel_rot, ha='right')
    else:
        ax1.set_xticklabels(limited_labels)
    if input_percent:
        ax1.yaxis.set_major_formatter(PercentFormatter(xmax=1.0))

    # adjust the second axis if cumplot=True
    if cumplot:
        yt = [str(int(it)) + r'%' for it in ax2.get_yticks()]
        ax2.set_yticklabels(yt)
    #     ax2.legend()
    # ax1.legend()
    fig.legend(loc=4, ncol=1)

    if grad is not None:
        grad(bars, rot90)
    # add data on top of bar
    for p in ax1.patches:
        width, height = p.get_width(), p.get_height()
        ax1.annotate('{:.2%}'.format(height), (p.get_x() - 0.3 * width, height + total_data / 10),
                     annotation_clip=False)
    # set y-axis tick numbers
    # ymin, ymax = ax1.get_ylim()
    # ax1.set_yticks(np.round(np.linspace(ymin, ymax, 6), 2))
    if cumplot:
        return fig, ax1, ax2, df_sorted
    else:
        return fig, ax1, df_sorted


def chart_gen(df_plot, plot_kw=None, title=None, annotation=None, grid=True, yaxis_format='percent', grad=None,
              ticksfontbold=True):
    if plot_kw is None:
        plot_kw = dict()
    last_idx = df_plot.last_valid_index()
    if last_idx == 'Baseline':
        y_ign_bse = True
    else:
        y_ign_bse = False
    ax = df_plot.plot(**plot_kw)
    ax.set_title(title, {'fontsize': 12, 'fontweight': 'bold'})
    handles, labels = ax.get_legend_handles_labels()
    # remove line legend if line axis input
    if plot_kw.get('ax') is not None:
        handles, labels = split_list_right(handles), split_list_right(labels)
    # reversed lables align legend with stacked bar
    ax.legend(list(reversed(handles)), list(reversed(labels)), bbox_to_anchor=(1.0, 0.5), loc="center left",
              labelspacing=1)
    if yaxis_format == 'percent':
        ax.yaxis.set_major_formatter(PercentFormatter(xmax=1.0))  # set y ax percent format
    # set x, y tick direction
    minor_locator = AutoMinorLocator(2)
    ax.xaxis.set_minor_locator(minor_locator)
    ax.xaxis.set_tick_params(bottom=False, top=False, which='major')
    ax.xaxis.set_tick_params(direction='in', which='minor', length=5)
    ax.yaxis.set_tick_params(direction='in', which='major')
    # x axis label
    ax.xaxis.label.set_visible(False)
    if ticksfontbold:
        for label in ax.get_xticklabels():
            label.set_fontweight('bold')
        # y axis label
        for label in ax.get_yticklabels():
            label.set_fontweight('bold')
    if grid:
        ax.grid(True, which='major', axis='y')
        ax.set_axisbelow(True)
    bottom, top = ax.get_ylim()
    # for pareto monthly summary chart
    if plot_kw['kind'] == 'line':
        plt.xticks(np.arange(len(df_plot)), df_plot.index.tolist(), rotation=90)
        top *= 1.5
        ax.set_ylim(top=top)

    # highlight Baseline with red dashed rectangle

    if y_ign_bse:
        p = ax.patches[-1]
        width, height = p.get_width(), p.get_height()
        x, y = p.get_x(), p.get_y()
        rect = patches.Rectangle((x - (0.3 * width), -0.3 * top), width * 1.5, 1.31 * top, linewidth=2,
                                 edgecolor='r', facecolor='none', linestyle='dashed', clip_on=False)
        ax.add_patch(rect)

    if annotation is not None:
        if annotation == 'top':
            for p in ax.patches[:-2] if y_ign_bse else ax.patches:
                width, height = p.get_width(), p.get_height()
                ax.annotate('{}'.format(height), (p.get_x() - 0.2 * width, height + top / 100), fontsize=12,
                            fontweight='bold', annotation_clip=False)
        elif annotation == 'mid':
            for p in ax.patches[:-1] if y_ign_bse else ax.patches:
                width, height = p.get_width(), p.get_height()
                x, y = p.get_x(), p.get_y()
                ax.annotate('{:.1%}'.format(height), (x + width, y + height / 2), ha='center', va='center',
                            fontsize=12, fontweight='bold', annotation_clip=False)
        elif annotation == 'top2':
            for p in ax.patches:
                width, height = p.get_width(), p.get_height()
                x, y = p.get_x(), p.get_y()
                ax.annotate('{}'.format(height), (x + 0.25 * width, height + top / 100), fontsize=12,
                            annotation_clip=False)

    if grad is not None:
        bars = ax.containers[0]
        grad(bars, rot90=True)

    ax.set_ylim(bottom=0)
    return ax


def insp_pareto(df_input, side, part, rot90=False, return_df_sorted_data=False):
    """

    :param df_input: multi index column df with column name 'FSV' and 'BSV'.
    :param rot90: default False, change gradient direction
    :param side: 'BSV' or 'FSV'
    :param return_df_sorted_data: default False, return sorted pareto df
    :return:
    """
    # if df_input.at[df_input.last_valid_index(), ('DATETIME', 'week')] == 'Baseline':
    df = df_input.copy()

    if chk_last_row(df, ('DATETIME', 'week'), 'Baseline'):
        df.drop(df.last_valid_index(), inplace=True)
    # else:
    #     df_input = df
    data = df.tail(1)[side].loc[[df.last_valid_index()]]

    fig, ax1, ax2, df_sorted_data = pareto(data, data_kw={'label': 'Yield loss%', 'width': 0.5,
                                                          'color': default_grad['midcolor'],
                                                          'linewidth': 1, 'edgecolor': 'k'},
                                           line_kw={'color': '#000080', 'marker': 'D', 'markersize': 5,
                                                    'label': 'Accumulated%'},
                                           not_others_num=5, xticklabel_rot=45, grad=gradientbars, rot90=rot90)
    plt.title(f'{part} {side} Yield Loss Pareto', fontsize=12)
    if return_df_sorted_data:
        return plt, df_sorted_data
    return plt
    # plt.show()


def stacked_bar_chart(df_input, side, part, title=None):
    df_plot = df_input.set_index(('DATETIME', 'week'))[side]
    if title is None:
        title = f'{part} {side} Yield Loss Trend'
    # add connection line between bar
    df_line = df_plot.iloc[np.arange(len(df_plot)).repeat(2)]
    df_line.reset_index(drop=True, inplace=True)
    df_line.index = df_line.index / 2 - bar_width / 2  # minus half bar width
    ax1 = df_line.plot(kind='line', stacked=True, color='k', linewidth=0.5)
    plot_kw = {'kind': 'bar', 'stacked': True, 'ax': ax1, 'figsize': (cm2inch(22), cm2inch(11)), 'fontsize': 12,
               'edgecolor': 'k', 'linewidth': 1, 'color': colordict[side], 'width': bar_width}
    ax2 = chart_gen(df_plot, plot_kw, title=title)
    return plt


def line_chart(df_input: pd.DataFrame, side: str, title: str, sorted_data: pd.DataFrame):
    # pareto monthly summary in presentation
    sorted_data.drop('Others', axis=1, inplace=True)
    df_plot = df_input[side].loc[:, sorted_data.columns]
    plot_kw = {'kind': 'line', 'stacked': False, 'figsize': (cm2inch(22), cm2inch(11)), 'fontsize': 12,
               'linewidth': 1, 'style': '.-', 'markersize': 12, 'clip_on': False}
    ax2 = chart_gen(df_plot, plot_kw, title=title)
    return plt


def weekly_yield(df_input, part):
    y_ign_bse = False
    title = f'{part} Weekly Yield Trend'
    df = df_input.set_index(('DATETIME', 'week'))
    last_idx = df.last_valid_index()
    if last_idx == 'Baseline':
        y_ign_bse = True
    else:
        df.rename(index={last_idx: '      ' + last_idx}, inplace=True)  # space prevent legend cover the plot
    # deduplicate index by add number ' ' in front of duplicate index by count
    s = pd.Series([' ' for n in range(len(df))])
    df.index = s.str.repeat(repeats=df.groupby(level=0).cumcount()) + df.index

    df_line = df['SUMMARY']['Yield']
    df_plot = df['SUMMARY']["Q'ty"]
    if y_ign_bse:
        df_temp = df_plot.drop(last_idx)
        ax1 = df_temp.plot()
        ylim = ax1.get_ylim()
        ax1.clear()
    plot_kw = {'kind': 'bar', 'figsize': (cm2inch(22), cm2inch(11)), 'fontsize': 12,
               'edgecolor': 'k', 'linewidth': 1, 'ylim': ylim if y_ign_bse else None,
               'width': bar_width}
    ax1 = chart_gen(df_plot, plot_kw, title=title, yaxis_format="Q'ty", grid=False, grad=gradientbars, annotation='top')

    ax1.get_legend().remove()
    ax1.set_ylabel("Q'ty", fontsize=12, fontweight='bold')
    ax2 = ax1.twinx()  # x tick will mismatch if df x axis has duplicate value

    # print(plt.xticks())
    ax2.plot(df_line, color='#000080', linewidth=1.5, marker='D', markersize=5)
    # print(plt.xticks())

    ax2.set_ylim([0.6, 1])
    ax2.set_ylabel('Yield', fontsize=12, fontweight='bold')
    ax2.yaxis.set_tick_params(direction='in', which='major')
    ax2.yaxis.set_major_formatter(PercentFormatter(xmax=1.0))  # set y ax percent format

    for label in ax2.get_yticklabels():
        label.set_fontweight('bold')
    bottom, top = ax2.get_ylim()
    line = ax2.lines[0].get_data()
    for x, y in zip(*line):
        ax2.annotate('{:.1%}'.format(y), (x, y + top / 20), fontsize=12, rotation=45, annotation_clip=False)
    fig = plt.gcf()
    leg = fig.legend(ncol=2, prop={'size': 12, 'weight': 'bold'}, loc=8)
    leg.legendHandles[0].set_color(default_grad['midcolor'])
    return plt
    # plt.show()


def weekly_yield_loss(df_input, part, color=None):
    if color is None:
        color = bsvcolors
    df = df_input.set_index(('DATETIME', 'week'))['SUMMARY']
    last_idx = df.last_valid_index()
    if not last_idx == 'Baseline':
        df.rename(index={last_idx: '      ' + last_idx}, inplace=True)  # prevent legend from cover the plot
    df_plot = df.drop(df.columns[[0, 1, 2]], axis=1)
    title = f'{part} Weekly Final Yield Loss Breakdown'
    plot_kw = {'kind': 'bar', 'stacked': True, 'figsize': (cm2inch(22), cm2inch(11)), 'fontsize': 12,
               'edgecolor': 'k', 'linewidth': 1, 'color': color, 'width': bar_width}
    ax1 = chart_gen(df_plot, plot_kw, title=title, grid=False, annotation='mid')
    ax1.get_legend().remove()
    ax1.set_ylabel("Yield loss %", fontsize=12, fontweight='bold')
    fig = plt.gcf()
    fig.legend(loc=8, ncol=3, prop={'size': 12, 'weight': 'bold'})
    return plt
    # plt.show()


def box_plots(df_input, part):
    fig, ax = plt.subplots(4, figsize=(cm2inch(22), cm2inch(44)))
    fs = 10  # fontsize
    # data filtering
    df = df_input.loc[:, df_input.columns.get_level_values(1).isin(
        {"week", 'mean', 'med', 'iqr', 'q1', 'q3', 'cilo', 'cihi', 'whislo', 'whishi', 'fliers'})]
    df = df.dropna()
    boxprops = dict(linestyle='-', linewidth=2)
    flierprops = dict(marker='x', markersize=8)
    meanpointprops = dict(marker='+', markersize=8)
    medianprops = dict(linestyle='-', linewidth=1.5)
    for i, p in enumerate(['OVERALL_BOX_PLOT', 'FSV_BOX_PLOT', 'BSV_BOX_PLOT', "CP_BOX_PLOT"]):
        stats = []
        for week, mean, med, iqr, q1, q3, cilo, cihi, whislo, whishi, fliers in df.loc[:, ["DATETIME", p]].itertuples(
                index=False):
            fliers = fliers.strip('[]')
            # fliers = fliers.strip('[')
            outliers = np.fromstring(fliers, sep=' ')

            stats.append(
                {'label': week, 'mean': mean, 'med': med, 'iqr': iqr, 'q1': q1, 'q3': q3, 'cilo': cilo, 'cihi': cihi,
                 'whislo': whislo, 'whishi': whishi, 'fliers': outliers})

        bplot = ax[i].bxp(stats, patch_artist=True, showmeans=True, flierprops=flierprops, meanprops=meanpointprops,
                          boxprops=boxprops, medianprops=medianprops)
        if i:  # yield
            s = ''
        else:  # yield loss
            s = ' loss'
        ax[i].set_title(f"{part} yield{s} {p.replace('_', ' ')}", fontsize=fs)
        ax[i].yaxis.set_major_formatter(PercentFormatter(xmax=1.0, decimals=2))
        ax[i].yaxis.grid(True)
        for patch in bplot['boxes']:
            patch.set_facecolor('gainsboro')
    # plt.show(block=True)
    # plt.interactive(False)
    return plt


def weekly_ooc():
    correction = {"Q'ty": 'Total', 'pcs': 'OOC'}
    df_input = get_week_ooc()
    idxmax = df_input.last_valid_index()
    year = df_input.at[idxmax, 'shipping_date'][-6:-4]
    week = df_input.at[idxmax, 'week']
    title = 'Y{}{} OOC rate'.format(year, week)
    df_plot = df_input[['pcs', "Q'ty"]].astype(int)
    df_plot.columns = [correction[n] for n in df_plot.columns]
    plot_kw = {'kind': 'bar', 'figsize': (cm2inch(12.65), cm2inch(7.65)), 'fontsize': 12, 'rot': 0,
               'edgecolor': 'k', 'linewidth': 0.5, 'width': bar_width, 'color': [c('#FF9895'), c('#4286D7')]}
    ax1 = chart_gen(df_plot, plot_kw, title=title, yaxis_format="Q'ty", annotation='top2', ticksfontbold=False)
    h, k = ax1.get_legend_handles_labels()
    ax1.get_legend().remove()
    df_dot = (df_plot['OOC'] / df_plot['Total']).fillna(0)
    ax2 = ax1.twinx()
    ax2.plot(df_dot, color='#90B546', linewidth=0, marker="^", markersize=7, label='OOC rate')

    ax2.yaxis.set_tick_params(direction='in', which='major')
    ax2.yaxis.set_major_formatter(PercentFormatter(xmax=1.0))
    ax2.set_ylim(bottom=0)  # fix ooc rate bottom 0.00%
    bottom, top = ax2.get_ylim()
    i, j = ax2.get_legend_handles_labels()
    line = ax2.lines[0].get_data()
    for x, y in zip(*line):
        ax2.annotate('  {:.1%}'.format(y), (x, y - top / 40), fontsize=12, annotation_clip=False)
    ax2.legend(h + i, k + j, loc='center left', bbox_to_anchor=(1.15, 0.5))
    plt.tight_layout()
    # plt.show()
    return plt, df_input


def set_bar_line_yrange(df_input):
    result = dict()
    bar_max = df_input["SUMMARY"]['Yield_loss'].max()
    line_max = df_input['SUMMARY']['Yield'].max()
    if round(line_max, 1) > 0.8:
        line_ymax = 1
        line_ymin = 0.6
    else:
        line_ymax = 0.9
        line_ymin = 0.5
    if round(bar_max, 1) < 0.2:
        bar_ymax = 0.2
        bar_ymin = 0
    else:
        bar_ymax = 0.5
        bar_ymin = 0
    result['line'] = [line_ymin, line_ymax]
    result['bar'] = [bar_ymin, bar_ymax]
    return result


def monthly_stacked_bar(df_input):
    last_idx = df_input.last_valid_index()
    df_input.rename(index={last_idx: '      ' + last_idx}, inplace=True)
    df_plot = df_input["SUMMARY"][['CP', 'BSV', 'FSV']]
    rename_dict = {grp: grp + ' Yield Loss' for grp in ['CP', 'BSV', 'FSV']}
    df_plot.rename(columns=rename_dict, inplace=True)
    df_line = df_input['SUMMARY']['Yield']
    plot_kw = {'kind': 'bar', 'stacked': True, 'figsize': (cm2inch(16.27), cm2inch(10)), 'fontsize': 12,
               'width': bar_width, 'color': monthly_colors}
    ax1 = chart_gen(df_plot, plot_kw)
    ax1.get_legend().remove()
    ax1.set_ylabel("Yield loss", fontsize=12, fontweight='bold')
    ax2 = ax1.twinx()
    ax2.plot(df_line, color='r', linewidth=1.5, marker='^', markersize=5)
    y_range = set_bar_line_yrange(df_input)
    ax1.set_ylim(y_range['bar'])
    ax2.set_ylim(y_range['line'])
    ax2.set_ylabel('Yield', fontsize=12, fontweight='bold')
    ax2.yaxis.set_tick_params(direction='in', which='major')
    ax2.yaxis.set_major_formatter(PercentFormatter(xmax=1.0))  # set y ax percent format
    for label in ax2.get_yticklabels():
        label.set_fontweight('bold')
    bottom, top = ax2.get_ylim()
    line = ax2.lines[0].get_data()
    for x, y in zip(*line):
        if np.isnan(y):
            continue
        ax2.annotate('{:.1%}'.format(y), (x, y + top / 20), fontsize=12, rotation=45, annotation_clip=False)
    fig = plt.gcf()
    fig.legend(ncol=4, prop={'size': 11, 'weight': 'bold'}, loc=8)
    return plt


def img_setup(plot, image_data, show):
    figure = plot.gcf()
    # add border around the figure
    figure.patches.extend([patches.Rectangle(xy=(0, 0.001), width=0.999, height=1, linewidth=2,
                                             transform=figure.transFigure, fill=False, edgecolor='#808080',
                                             figure=figure)])
    plot.tight_layout()
    plot.savefig(image_data, format="png")
    if show:
        plt.show(block=True)
        plt.interactive(False)
    plot.clf()
    plot.close()
    image_data.seek(0)
    return Image(image_data)


class ChartIOGen:

    def __init__(self, df=pd.DataFrame(), part=""):
        self.df = df
        self.part = part

    def weekly_defect_data(self, kind, site=None, show=False, return_sorted_data=False):
        image_data = BytesIO()
        insp = {'pareto': insp_pareto, 'stacked_bar_chart': stacked_bar_chart}
        yields = {'yield': weekly_yield, 'yield_loss': weekly_yield_loss, 'box_plot': box_plots}
        if site is None:
            gen_func = yields[kind]
            plot = gen_func(self.df, part=self.part)
        else:
            gen_func = insp[kind]
            if return_sorted_data and kind == 'pareto':
                plot, sorted_data = gen_func(self.df, site, part=self.part, return_df_sorted_data=return_sorted_data)
                return img_setup(plot, image_data, show), sorted_data
            plot = gen_func(self.df, site, part=self.part)
        return img_setup(plot, image_data, show)

    def monthly_data(self, show=False, monthly_pareto_summary=False, side=None, title=None, sorted_data=None):
        """
        :param show:show
        :param monthly_pareto_summary: generate side pareto monthly summary chart
        :param side: side
        :param title: title
        :param sorted_data:sorted_data
        :return: img
        """

        cross_year = not chk_month_order(self.df)

        if cross_year:
            df_month = self.df.tail(1)
        else:
            df_month = self.df
        df_month = monthly_data_fill(df_month)
        image_data = BytesIO()
        if monthly_pareto_summary:
            plot = line_chart(df_month, side, title=title, sorted_data=sorted_data)
        else:
            plot = monthly_stacked_bar(df_month)
        return img_setup(plot, image_data, show)

    # def last_two_month_diff(self):

    @staticmethod
    def weekly_ooc(show=False):
        image_data = BytesIO()
        plot, df_info = weekly_ooc()
        return img_setup(plot, image_data, show), df_info


def main():
    pass


if __name__ == '__main__':
    # main5()
    # main1()
    # main2()
    # main3()
    # main4()
    # main6()
    main()
