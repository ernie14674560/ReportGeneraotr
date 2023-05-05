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
from wm_utils import flat_find_and_flip


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
                 die_size,
                 wafer_detail: dict,  # add by Ernest
                 center_xy=(0, 0),
                 dia=150,
                 edge_excl=5,
                 flat_excl=5,
                 data_type=wm_const.DataType.CONTINUOUS,
                 high_color=wm_const.wm_HIGH_COLOR,
                 low_color=wm_const.wm_LOW_COLOR,
                 plot_range=None,
                 plot_die_centers=False,
                 show_die_gridlines=True,
                 title="Wafer Map Phoenix",
                 discrete_legend_values=None,
                 discrete_legend_colors=None,
                 icon=None
                 ):
        import wx

        self.app = wx.App()
        atexit.register(disable_asserts)

        self.wafer_info = wm_info.WaferInfo(die_size,
                                            center_xy,
                                            dia,
                                            edge_excl,
                                            flat_excl,
                                            )
        self.wafer_detail = wafer_detail  # add by Ernest
        df_map = self.wafer_detail['df_map']  # add by Ernest
        self.flip_y = flat_find_and_flip(df_map)
        # flat_find_and_flip(df_map)
        xyd = list(df_map.itertuples(index=False))  # add by Ernest
        self.xyd = xyd
        self.data_type = data_type
        self.high_color = high_color
        self.low_color = low_color
        self.plot_range = plot_range
        self.plot_die_centers = plot_die_centers
        self.show_die_gridlines = show_die_gridlines
        self.title = title
        self.discrete_legend_values = discrete_legend_values
        self.discrete_legend_colors = discrete_legend_colors
        self.frame = wm_frame.WaferMapWindow(self.title,
                                             self.xyd,
                                             self.wafer_info,
                                             self.wafer_detail,
                                             parent=self,
                                             flip_y=self.flip_y,
                                             data_type=self.data_type,
                                             high_color=self.high_color,
                                             low_color=self.low_color,
                                             #                                      high_color=wx.Colour(255, 0, 0),
                                             #                                      low_color=wx.Colour(0, 0, 255),
                                             plot_range=self.plot_range,
                                             size=(600, 500),
                                             plot_die_centers=self.plot_die_centers,
                                             show_die_gridlines=self.show_die_gridlines,
                                             discrete_legend_values=self.discrete_legend_values,
                                             discrete_legend_colors=self.discrete_legend_colors
                                             )
        self.frame.Maximize()  # modify by Ernest
        if icon is not None:
            self.frame.SetIcon(wx.Icon(icon))
        self.frame.Show()
        self.frame.on_zoom_fit(event='')

        self.app.MainLoop()


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
