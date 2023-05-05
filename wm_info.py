# -*- coding: utf-8 -*-
"""
The :class:`wafer_map.wm_info.WaferInfo` class.
"""
import wm_constants as wm_const
from wm_utils import flat_find_and_flip


class WaferInfo(object):
    """
    Contains the wafer information.

    Parameters
    ----------
    die_size : tuple
        The die size in mm as a ``(width, height)`` tuple.
    center_xy : tuple
        The center grid coordinate as a ``(x_grid, y_grid)`` tuple.
    dia : float, optional
        The wafer diameter in mm. Defaults to `150`.
    edge_excl : float, optional
        The distance in mm from the edge of the wafer that should be
        considered bad die. Defaults to 5mm.
    flat_excl : float, optional
        The distance in mm from the wafer flat that should be
        considered bad die. Defaults to 5mm.
    wafer_detail:dict
        {df_map:.., tab_title:..}
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
    """

    def __init__(self, die_size, center_xy, wafer_detail,
                 dia=150,
                 edge_excl=5,
                 flat_excl=5,
                 data_type=wm_const.DataType.CONTINUOUS,
                 high_color=wm_const.wm_HIGH_COLOR,
                 low_color=wm_const.wm_LOW_COLOR,
                 plot_range=None,
                 discrete_legend_values=None,
                 discrete_legend_colors=None,
                 plot_die_centers=False,
                 show_die_gridlines=True,
                 show_wafer_objects=False
                 ):
        self.die_size = die_size
        self.center_xy = center_xy
        self.dia = dia
        self.edge_excl = edge_excl
        self.flat_excl = flat_excl
        self.wafer_detail = wafer_detail
        self.df_map = self.wafer_detail['df_map']
        self.flip_y = flat_find_and_flip(self.df_map)
        xyd = list(self.df_map.itertuples(index=False))  # add by Ernest
        self.xyd = xyd
        # backwards compatibility
        if isinstance(data_type, str):
            data_type = wm_const.DataType(data_type)
        self.data_type = data_type
        self.high_color = high_color
        self.low_color = low_color
        self.plot_range = plot_range
        self.discrete_legend_values = discrete_legend_values
        self.discrete_legend_colors = discrete_legend_colors
        self.tab_name = self.wafer_detail['tab_title']
        self.plot_die_centers = plot_die_centers
        self.show_die_gridlines = show_die_gridlines
        self.show_wafer_objects = show_wafer_objects

    def __str__(self):
        string = """
Wafer Dia: {}mm
Die Size: {}
Grid Center XY: {}
Edge Excl: {}
Flat Excl: {}
"""
        return string.format(self.dia,
                             self.die_size,
                             self.center_xy,
                             self.edge_excl,
                             self.flat_excl,
                             )


def main():
    """Run when called as a module."""
    raise RuntimeError("This module is not meant to be run by itself.")


if __name__ == "__main__":
    main()
