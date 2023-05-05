# -*- coding: utf-8 -*-
"""
A self-contained Window for a Wafer Map.
"""
# ---------------------------------------------------------------------------
### Imports
# ---------------------------------------------------------------------------
# Standard Library
from __future__ import absolute_import, division, print_function, unicode_literals

import atexit

# Third-Party
# from . import gen_fake_data
import wm_constants as wm_const
# Package / Application
import wm_frame
import wm_info


# import wx


class WaferMapApp(object):
    """
    A self-contained Window for a Wafer Map.

    Parameters
    ----------
    xyd : list of 3-tuples
        The data to plot.
    die_size : tuple
        The die size in mm as a ``(width, height)`` tuple.
    center_xy : tuple, optional
        The center grid coordinate as a ``(x_grid, y_grid)`` tuple.
    dia : float, optional
        The wafer diameter in mm. Defaults to `150`.
    edge_excl : float, optional
        The distance in mm from the edge of the wafer that should be
        considered bad die. Defaults to 5mm.
    flat_excl : float, optional
        The distance in mm from the wafer flat that should be
        considered bad die. Defaults to 5mm.
    data_type : wm_constants.DataType or str, optional
        The type of data to plot. Must be one of `continuous` or `discrete`.
    high_color : :class:`wx.Colour`, optional
        The color to display if a value is above the plot range. Defaults
        to `wm_constants.wm_HIGH_COLOR`.
    low_color : :class:`wx.Colour`, optional
        The color to display if a value is below the plot range. Defaults
        to `wm_constants.wm_LOW_COLOR`.
    plot_range : tuple, optional
        The plot range to display. If ``None``, then auto-ranges. Defaults
        to auto-ranging.
    plot_die_centers : bool, optional
        If ``True``, display small red circles denoting the die centers.
        Defaults to ``False``.
    show_die_gridlines : bool, optional
        If ``True``, displayes gridlines along the die edges. Defaults to
        ``True``.
    discrete_legend_values : list, optional
        A list of strings for die bins. Every data value in ``xyd`` must
        be in this list. This will define the legend order. Only used when
        ``data_type`` is ``discrete``.
    discrete_legend_colors : list, optional
        A list of [(R, G, B), ...] color for die bins. Every data value in ``xyd`` must
        be in this list. This will define the legend order color. Only used when
        ``data_type`` is ``discrete``.
    """

    def __init__(self,
                 wafer_list: list,
                 center_xy=(0, 0),
                 dia=150,
                 edge_excl=5,
                 flat_excl=5,
                 high_color=wm_const.wm_HIGH_COLOR,
                 low_color=wm_const.wm_LOW_COLOR,
                 plot_die_centers=False,
                 show_die_gridlines=True,
                 # die_size,
                 # wafer_detail: dict,  # add by Ernest
                 # data_type=wm_const.DataType.CONTINUOUS,
                 # plot_range=None,
                 # discrete_legend_values=None,
                 # discrete_legend_colors=None,
                 icon='icon2.png',
                 show_wafer_objects=False,
                 title="Wafer Map Phoenix",
                 auto_snapshot=False,
                 # first=True
                 ):
        import wx
        atexit.register(disable_asserts)

        app = wx.App()
        # app = wx.App()
        self.wafer_infos = []
        for wafer in wafer_list:
            wafer_info = wm_info.WaferInfo(**wafer,
                                           center_xy=center_xy,
                                           dia=dia,
                                           edge_excl=edge_excl,
                                           flat_excl=flat_excl,
                                           high_color=high_color,
                                           low_color=low_color,
                                           plot_die_centers=plot_die_centers,
                                           show_die_gridlines=show_die_gridlines,
                                           show_wafer_objects=show_wafer_objects
                                           )
            self.wafer_infos.append(wafer_info)
        self.plot_die_centers = plot_die_centers
        self.show_die_gridlines = show_die_gridlines
        self.title = title

        self.frame = wm_frame.WaferMapWindow(self.title,

                                             self.wafer_infos,

                                             size=(600, 500),
                                             show_wafer_objects=show_wafer_objects
                                             )
        self.frame.Maximize()  # modify by Ernest
        if icon is not None:
            self.frame.SetIcon(wx.Icon(icon))
        self.frame.Show()

        self.frame.on_zoom_fit(event='')
        if auto_snapshot:
            wx.CallLater(2000, self.auto_snapshot)
        app.MainLoop()
        del app

    def Close(self):
        import wx
        wx.CallAfter(self.frame.Close)

    def auto_snapshot(self):
        import wx
        frame = self.frame
        frame.on_export_snapshot('', True, r'C:\Users\ernie.chen\Desktop\ReportGen ver2\AP197 Rin')
        wx.CallAfter(self.frame.Close)


def disable_asserts():
    import wx
    wx.DisableAsserts()


def main():
    # """Run when called as a module."""
    # wafer_info, xyd = gen_fake_data.generate_fake_data(die_x=5.43,
    #                                                    die_y=6.3,
    #                                                    dia=150,
    #                                                    edge_excl=4.5,
    #                                                    flat_excl=4.5,
    #                                                    x_offset=0,
    #                                                    y_offset=0.5,
    #                                                    grid_center=(29, 21.5),
    #                                                    )
    #
    # import random
    # bins = ["Bin1", "Bin1", "Bin1", "Bin2", "Dragons", "Bin1", "Bin2"]
    # discrete_xyd = [(_x, _y, random.choice(bins))
    #                 for _x, _y, _
    #                 in xyd]
    #
    # discrete = False
    # dtype = wm_const.DataType.CONTINUOUS
    #
    # #    discrete = True         # uncomment this line to use discrete data
    # if discrete:
    #     xyd = discrete_xyd
    #     dtype = wm_const.DataType.DISCRETE
    #
    # WaferMapApp(xyd,
    #             wafer_info.die_size,
    #             wafer_info.center_xy,
    #             wafer_info.dia,
    #             wafer_info.edge_excl,
    #             wafer_info.flat_excl,
    #             data_type=dtype,
    #             #                plot_range=(0.0, 75.0**2),
    #             plot_die_centers=True,
    #             )
    pass


if __name__ == "__main__":
    main()
#    pass
