# -*- coding: utf-8 -*-
# pylint: disable=E1101
#   E1101 = Module X has no Y member
"""
This is the main window of the Wafer Map application.
"""
# ---------------------------------------------------------------------------
### Imports
# ---------------------------------------------------------------------------
# Standard Library
from __future__ import absolute_import, division, print_function, unicode_literals

import os
import subprocess

# Third-Party
import wx

import wm_constants as wm_const
# Package / Application
import wm_core
import wm_utils

# import ctypes
# def find_window(parent, names):
#     if not names:
#         return parent
#     name = str(names[0])
#     child = 0
#     while True:
#         child = ctypes.windll.user32.FindWindowExW(parent, child, name, 0)
#         if not child:
#             return 0
#         result = find_window(child, names[1:])
#         if result:
#             return result
#
#
# def find_handle():
#     desktop = ctypes.windll.user32.GetDesktopWindow()
#     handle = 0
#     handle = handle or find_window(desktop, ['Progman', 'SHELLDLL_DefView', 'SysListView32'])
#     handle = handle or find_window(desktop, ['WorkerW', 'SHELLDLL_DefView', 'SysListView32'])
#     return handle

# current_uid = 0
#
#
# def find_handle():
#     handle = current_uid + 1
#     return handle


class WaferMapWindow(wx.Frame):
    """
    This is the main window of the application.

    It contains the WaferMapPanel and the MenuBar.

    Although technically I don't need to have only 1 panel in the MainWindow,
    I can have multiple panels. But I think I'll stick with this for now.

    Parameters
    ----------
    title : str
        The title to display.
    xyd : list of 3-tuples
        The data to plot.
    wafer_info : :class:`wx_info.WaferInfo`
        The wafer information.
    size : tuple, optional
        The windows size in ``(width, height)``. Values must be ``int``s.
        Defaults to ``(800, 600)``.
    data_type : wm_constants.DataType or string, optional
        The type of data to plot. Must be one of `continuous` or `discrete`.
        Defaults to `continuous`.
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
                 title,
                 xyd,
                 wafer_info,
                 wafer_detail: dict,
                 size=(800, 600),
                 flip_y=False,
                 parent=None,
                 data_type=wm_const.DataType.CONTINUOUS,
                 high_color=wm_const.wm_HIGH_COLOR,
                 low_color=wm_const.wm_LOW_COLOR,
                 plot_range=None,
                 plot_die_centers=False,
                 show_die_gridlines=True,
                 show_wafer_objects=False,
                 discrete_legend_values=None,
                 discrete_legend_colors=None
                 ):
        # def __init__(self, parent=None, id=None, title=None, pos=None, size=None, style=None, name=None)
        # wx.Frame.__init__(self, parent, id=wx.ID_ANY, title=u"Encrypted Files", pos=wx.DefaultPosition,
        #                   size=wx.Size(600, 400), style=wx.DEFAULT_FRAME_STYLE | wx.TAB_TRAVERSAL)
        wx.Frame.__init__(self,
                          None,
                          id=wx.ID_ANY,
                          title=title,
                          size=size,
                          )
        # handle = find_handle()
        # self.AssociateHandle(handle)
        self.title = title
        self.wafer_detail = wafer_detail
        self.flip_y = flip_y
        self.xyd = xyd
        self.wafer_info = wafer_info
        # backwards compatability
        if isinstance(data_type, str):
            data_type = wm_const.DataType(data_type)
        self.data_type = data_type
        self.high_color = high_color
        self.low_color = low_color
        self.plot_range = plot_range
        self.plot_die_centers = plot_die_centers
        self.show_die_gridlines = show_die_gridlines
        self.show_wafer_objects = show_wafer_objects  # add by Ernest
        self.discrete_legend_values = discrete_legend_values
        self.discrete_legend_colors = discrete_legend_colors

        self._init_ui()

    def _init_ui(self):
        """Init the UI components."""
        # Create menu bar
        self.menu_bar = wx.MenuBar()

        self._create_menus()
        self._create_menu_items()
        self._add_menu_items()
        self._add_menus()
        self._bind_events()

        # Initialize default states
        if self.show_wafer_objects:
            self.mv_outline.Check()
            self.mv_crosshairs.Check()
        self.mv_diecenters.Check(self.plot_die_centers)
        self.mv_legend.Check()

        # Set the MenuBar and create a status bar (easy thanks to wx.Frame)
        self.SetMenuBar(self.menu_bar)
        self.CreateStatusBar()

        # Allows this module to be run by itself if needed.
        if __name__ == "__main__":
            self.panel = None
        else:
            self.panel = wm_core.WaferMapPanel(self,
                                               self.xyd,
                                               self.wafer_info,
                                               self.wafer_detail,
                                               flip_y=self.flip_y,
                                               data_type=self.data_type,
                                               high_color=self.high_color,
                                               low_color=self.low_color,
                                               plot_range=self.plot_range,
                                               plot_die_centers=self.plot_die_centers,
                                               show_die_gridlines=self.show_die_gridlines,
                                               show_wafer_objects=self.show_wafer_objects,
                                               discrete_legend_values=self.discrete_legend_values,
                                               discrete_legend_colors=self.discrete_legend_colors
                                               )

    # TODO: There's gotta be a more scalable way to make menu items
    #       and bind events... I'll run out of names if I have too many items.
    #       If I use numbers, as displayed in wxPython Demo, then things
    #       become confusing if I want to reorder things.
    def _create_menus(self):
        """Create each menu for the menu bar."""
        self.mfile = wx.Menu()
        self.medit = wx.Menu()
        self.mview = wx.Menu()
        self.mopts = wx.Menu()

    def _create_menu_items(self):
        """Create each item for each menu."""
        ### Menu: File (mf_) ###
        #        self.mf_new = wx.MenuItem(self.mfile,
        #                                  wx.ID_ANY,
        #                                  "&New\tCtrl+N",
        #                                  "TestItem")
        #        self.mf_open = wx.MenuItem(self.mfile,
        #                                   wx.ID_ANY,
        #                                   "&Open\tCtrl+O",
        #                                   "TestItem")
        self.mf_close = wx.MenuItem(self.mfile,
                                    wx.ID_ANY,
                                    "&Close\tCtrl+Q",
                                    "TestItem",
                                    )
        self.mf_export = wx.MenuItem(self.mfile,
                                     wx.ID_ANY,
                                     "&Export to excel\tCtrl+X",
                                     "Export to excel",
                                     )

        ### Menu: Edit (me_) ###
        self.me_redraw = wx.MenuItem(self.medit,
                                     wx.ID_ANY,
                                     "&Redraw",
                                     "Force Redraw",
                                     )

        ### Menu: View (mv_) ###
        self.mv_zoomfit = wx.MenuItem(self.mview,
                                      wx.ID_ANY,
                                      "Zoom &Fit\tHome",
                                      "Zoom to fit",
                                      )
        if self.show_wafer_objects:
            self.mv_crosshairs = wx.MenuItem(self.mview,
                                             wx.ID_ANY,
                                             "Crosshairs\tC",
                                             "Show or hide the crosshairs",
                                             wx.ITEM_CHECK,
                                             )
            self.mv_outline = wx.MenuItem(self.mview,
                                          wx.ID_ANY,
                                          "Wafer Outline\tO",
                                          "Show or hide the wafer outline",
                                          wx.ITEM_CHECK,
                                          )
        self.mv_diecenters = wx.MenuItem(self.mview,
                                         wx.ID_ANY,
                                         "Die Centers\tD",
                                         "Show or hide the die centers",
                                         wx.ITEM_CHECK,
                                         )
        self.mv_legend = wx.MenuItem(self.mview,
                                     wx.ID_ANY,
                                     "Legend\tL",
                                     "Show or hide the legend",
                                     wx.ITEM_CHECK,
                                     )

        # Menu: Options (mo_) ###
        self.mo_test = wx.MenuItem(self.mopts,
                                   wx.ID_ANY,
                                   "&Test",
                                   "Nothing",
                                   )
        self.mo_high_color = wx.MenuItem(self.mopts,
                                         wx.ID_ANY,
                                         "Set &High Color",
                                         "Choose the color for high values",
                                         )
        self.mo_low_color = wx.MenuItem(self.mopts,
                                        wx.ID_ANY,
                                        "Set &Low Color",
                                        "Choose the color for low values",
                                        )

    def _add_menu_items(self):
        """Append MenuItems to each menu."""
        #        self.mfile.Append(self.mf_new)
        #        self.mfile.Append(self.mf_open)
        self.mfile.Append(self.mf_close)
        self.mfile.Append(self.mf_export)
        self.medit.Append(self.me_redraw)
        #        self.medit.Append(self.me_test1)
        #        self.medit.Append(self.me_test2)

        self.mview.Append(self.mv_zoomfit)
        self.mview.AppendSeparator()
        if self.show_wafer_objects:
            self.mview.Append(self.mv_crosshairs)
            self.mview.Append(self.mv_outline)
        self.mview.Append(self.mv_diecenters)
        self.mview.Append(self.mv_legend)

        self.mopts.Append(self.mo_test)
        self.mopts.Append(self.mo_high_color)
        self.mopts.Append(self.mo_low_color)

    def _add_menus(self):
        """Append each menu to the menu bar."""
        self.menu_bar.Append(self.mfile, "&File")
        self.menu_bar.Append(self.medit, "&Edit")
        self.menu_bar.Append(self.mview, "&View")
        self.menu_bar.Append(self.mopts, "&Options")

    def _bind_events(self):
        """Bind events to varoius MenuItems."""
        self.Bind(wx.EVT_MENU, self.on_quit, self.mf_close)
        # self.Bind(wx.EVT_MENU, self.on_export_excel, self.mf_close)
        self.Bind(wx.EVT_MENU, self.on_export_excel, self.mf_export)
        self.Bind(wx.EVT_MENU, self.on_zoom_fit, self.mv_zoomfit)
        if self.show_wafer_objects:
            self.Bind(wx.EVT_MENU, self.on_toggle_crosshairs, self.mv_crosshairs)
            self.Bind(wx.EVT_MENU, self.on_toggle_outline, self.mv_outline)
        self.Bind(wx.EVT_MENU, self.on_toggle_diecenters, self.mv_diecenters)
        self.Bind(wx.EVT_MENU, self.on_toggle_legend, self.mv_legend)
        self.Bind(wx.EVT_MENU, self.on_change_high_color, self.mo_high_color)
        self.Bind(wx.EVT_MENU, self.on_change_low_color, self.mo_low_color)

        # If I define an ID to the menu item, then I can use that instead of
        #   and event source:
        # self.mo_test = wx.MenuItem(self.mopts, 402, "&Test", "Nothing")
        # self.Bind(wx.EVT_MENU, self.on_zoom_fit, id=402)

    def close(self):
        # self.Close(True)
        self.Destroy()

    def on_export_excel(self, event):
        if self.data_type == wm_const.DataType.DISCRETE:
            color_func = self.panel.legend.color_dict.get
        else:
            color_func = self.panel.legend.get_color

        with wx.FileDialog(self, "Save XLSX file",
                           wildcard="XLSX files (*.xlsx)|*.xlsx",
                           defaultFile=self.title,
                           defaultDir=os.path.dirname(os.getcwd()),
                           style=wx.FD_SAVE | wx.FD_OVERWRITE_PROMPT) as fileDialog:

            if fileDialog.ShowModal() == wx.ID_CANCEL:
                return  # the user changed their mind

            # save the current contents in the file
            pathname = fileDialog.GetPath()
            try:
                die_size = self.wafer_info.die_size
                die_size = tuple([x / die_size[0] * 5 for x in die_size])
                df_map = self.wafer_detail['df_map']
                wm_utils.export_to_excel(df_map, color_func, pathname, die_size)
            # with open(pathname, 'w') as file:
            #     self.doSaveData(file)

            except IOError:
                wx.LogError("Cannot save current data in file '%s'." % pathname)
            except Exception as e:
                wx.LogError(str(e))

            resp = wx.MessageBox('Will open file, continue?', 'Open file', wx.OK | wx.CANCEL)
            if resp == wx.OK:
                subprocess.Popen(["explorer", os.path.abspath(pathname)])

    def on_quit(self, event):
        """Action for the quit event."""
        self.Close(True)

    # TODO: I don't think I need a separate method for this
    def on_zoom_fit(self, event):
        """Call :meth:`wafer_map.wm_core.WaferMapPanel.zoom_fill()`."""
        print("Frame Event!")
        self.panel.zoom_fill()

    # TODO: I don't think I need a separate method for this
    def on_toggle_crosshairs(self, event):
        """Call :meth:`wafer_map.wm_core.WaferMapPanel.toggle_crosshairs()`."""
        self.panel.toggle_crosshairs()

    # TODO: I don't think I need a separate method for this
    def on_toggle_diecenters(self, event):
        """Call :meth:`wafer_map.wm_core.WaferMapPanel.toggle_crosshairs()`."""
        self.panel.toggle_die_centers()

    # TODO: I don't think I need a separate method for this
    def on_toggle_outline(self, event):
        """Call the WaferMapPanel.toggle_outline() method."""
        self.panel.toggle_outline()

    # TODO: I don't think I need a separate method for this
    #       However if I don't use these then I have to
    #           1) instance self.panel at the start of __init__
    #           2) make it so that self.panel.toggle_legend accepts arg: event
    def on_toggle_legend(self, event):
        """Call the WaferMapPanel.toggle_legend() method."""
        self.panel.toggle_legend()

    # TODO: See the 'and' in the docstring? Means I need a separate method!
    def on_change_high_color(self, event):
        """Change the high color and refresh display."""
        print("High color menu item clicked!")
        cd = wx.ColourDialog(self)
        cd.GetColourData().SetChooseFull(True)

        if cd.ShowModal() == wx.ID_OK:
            new_color = cd.GetColourData().Colour
            print("The color {} was chosen!".format(new_color))
            self.panel.on_color_change({'high': new_color, 'low': None})
            self.panel.Refresh()
        else:
            print("no color chosen :-(")
        cd.Destroy()

    # TODO: See the 'and' in the docstring? Means I need a separate method!
    def on_change_low_color(self, event):
        """Change the low color and refresh display."""
        print("Low Color menu item clicked!")
        cd = wx.ColourDialog(self)
        cd.GetColourData().SetChooseFull(True)

        if cd.ShowModal() == wx.ID_OK:
            new_color = cd.GetColourData().Colour
            print("The color {} was chosen!".format(new_color))
            self.panel.on_color_change({'high': None, 'low': new_color})
            self.panel.Refresh()
        else:
            print("no color chosen :-(")
        cd.Destroy()


def main():
    """Run when called as a module."""
    app = wx.App()
    frame = WaferMapWindow("Testing", [], None)
    frame.Show()
    app.MainLoop()


if __name__ == "__main__":
    main()
