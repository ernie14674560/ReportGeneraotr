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
import time

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

class MyNotebook(wx.Notebook):
    def __getitem__(self, index):
        """ More pythonic way to get a specific page, also useful for iterating
            over all pages, e.g: for page in notebook: ...
        """
        if index < self.GetPageCount():
            return self.GetPage(index)
        else:
            raise IndexError


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
                 # xyd,
                 wafer_infos,
                 # wafer_detail: dict,
                 size=(800, 600),
                 # flip_y=False,
                 # parent=None,
                 # data_type=wm_const.DataType.CONTINUOUS,
                 # high_color=wm_const.wm_HIGH_COLOR,
                 # low_color=wm_const.wm_LOW_COLOR,
                 # plot_range=None,
                 plot_die_centers=False,
                 show_die_gridlines=True,
                 show_wafer_objects=False,
                 # discrete_legend_values=None,
                 # discrete_legend_colors=None
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
        self.title = title
        self.wafer_infos = wafer_infos
        self.plot_die_centers = plot_die_centers
        self.show_die_gridlines = show_die_gridlines
        self.show_wafer_objects = show_wafer_objects  # add by Ernest
        self.capture_delay_time = 0.5  # 1 second
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
            pass
        else:
            self.p = wx.Panel(self)
            self.notebook = MyNotebook(self.p)
            for wafer_info in self.wafer_infos:
                page = wm_core.WaferMapPanel(parent=self.notebook,
                                             wafer_info=wafer_info
                                             )
                self.notebook.AddPage(page, wafer_info.tab_name)
            sizer = wx.BoxSizer()

            sizer.Add(self.notebook, 1, wx.EXPAND)

            self.p.SetSizer(sizer)

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
        self.mf_export_snapshot = wx.MenuItem(self.mfile,
                                              wx.ID_ANY,
                                              "&Export to image\tCtrl+Shift+X",
                                              "Export to image",
                                              )
        if len(self.wafer_infos) > 1:
            self.mf_export_all = wx.MenuItem(self.mfile,
                                             wx.ID_ANY,
                                             "&Export all pages to excel\tCtrl+S",
                                             "Export all pages to excel",
                                             )
            self.mf_export_all_snapshot = wx.MenuItem(self.mfile,
                                                      wx.ID_ANY,
                                                      "&Export all pages to images\tCtrl+Shift+S",
                                                      "Export all pages to images",
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
        self.mv_nexttab = wx.MenuItem(self.mview,
                                      wx.ID_ANY,
                                      "Next tab\tTab",
                                      "Next tab",
                                      )
        self.mv_previoustab = wx.MenuItem(self.mview,
                                          wx.ID_ANY,
                                          "Previous tab\tShift+Tab",
                                          "Previous tab",
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
        self.mo_delay_time = wx.MenuItem(self.mopts,
                                         wx.ID_ANY,
                                         "&Set capture image delay time",
                                         "Set capture image delay time",
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
        self.mfile.Append(self.mf_export_snapshot)
        if len(self.wafer_infos) > 1:
            self.mfile.Append(self.mf_export_all)
            self.mfile.Append(self.mf_export_all_snapshot)
        self.medit.Append(self.me_redraw)
        #        self.medit.Append(self.me_test1)
        #        self.medit.Append(self.me_test2)

        self.mview.Append(self.mv_zoomfit)
        self.mview.Append(self.mv_nexttab)
        self.mview.Append(self.mv_previoustab)
        self.mview.AppendSeparator()
        if self.show_wafer_objects:
            self.mview.Append(self.mv_crosshairs)
            self.mview.Append(self.mv_outline)
        self.mview.Append(self.mv_diecenters)
        self.mview.Append(self.mv_legend)

        self.mopts.Append(self.mo_delay_time)
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
        self.Bind(wx.EVT_MENU, self.on_export_snapshot, self.mf_export_snapshot)
        if len(self.wafer_infos) > 1:
            self.Bind(wx.EVT_MENU, self.on_multi_export_excel, self.mf_export_all)
            self.Bind(wx.EVT_MENU, self.on_multi_export_snapshoot, self.mf_export_all_snapshot)
        self.Bind(wx.EVT_MENU, self.on_zoom_fit, self.mv_zoomfit)
        self.Bind(wx.EVT_MENU, self.on_next_tab, self.mv_nexttab)
        self.Bind(wx.EVT_MENU, self.on_previous_tab, self.mv_previoustab)
        if self.show_wafer_objects:
            self.Bind(wx.EVT_MENU, self.on_toggle_crosshairs, self.mv_crosshairs)
            self.Bind(wx.EVT_MENU, self.on_toggle_outline, self.mv_outline)
        self.Bind(wx.EVT_MENU, self.on_toggle_diecenters, self.mv_diecenters)
        self.Bind(wx.EVT_MENU, self.on_toggle_legend, self.mv_legend)
        self.Bind(wx.EVT_MENU, self.on_change_high_color, self.mo_high_color)
        self.Bind(wx.EVT_MENU, self.on_change_low_color, self.mo_low_color)
        self.Bind(wx.EVT_MENU, self.on_change_delay_time, self.mo_delay_time)
        # If I define an ID to the menu item, then I can use that instead of
        #   and event source:
        # self.mo_delay_time = wx.MenuItem(self.mopts, 402, "&Test", "Nothing")
        # self.Bind(wx.EVT_MENU, self.on_zoom_fit, id=402)

    # def close(self):
    #     # self.Close(True)
    #     self.Destroy()

    # def onClose(self, event):
    #     """"""
    #     event.Skip()
    #     self.Destroy()
    def onExit(self, event):
        """"""
        # self.Close()
        print("QQ")
    @staticmethod
    def single_export(page, pathname, first=False):
        if page.data_type == wm_const.DataType.DISCRETE:
            color_func = page.legend.color_dict.get
        else:
            color_func = page.legend.get_color
        die_size = page.die_size
        die_size = tuple([x / die_size[0] * 5 for x in die_size])
        df_map = page.df_map

        pathname, df_data = wm_utils.export_to_excel(df_map=df_map, color_func=color_func, filename=pathname,
                                                     die_size=die_size, tab_title=page.tab_name, first=first,
                                                     plot_range=page.plot_range)
        return pathname, df_data
        # if return_writer:
        #     return pathname

    def on_multi_export_snapshoot(self, event):
        self.on_export_snapshot(event, multi=True)

    def export_snapshot(self, event, multi, pathname):
        self.Refresh()
        try:
            if multi:
                for i, page in enumerate(self.notebook):
                    self.notebook.SetSelection(i)
                    # page.Refresh()
                    wx.SafeYield()
                    time.sleep(self.capture_delay_time / 2)
                    # wx.CallLater(500, self.onTakeSnapShot, event, pathname, page)
                    self.onTakeSnapShot(event, pathname, page)
                    time.sleep(self.capture_delay_time / 2)
                    # self.onTakeSnapShot(event, pathname, page)
            else:
                page = self.notebook.CurrentPage
                self.onTakeSnapShot(event, pathname, page)
        except IOError:
            wx.LogError("Cannot save current data in directory '%s'." % pathname)
        except Exception as e:
            wx.LogError(str(e))

    def on_export_snapshot(self, event, multi=False, pathname=None):
        if pathname is None:
            with wx.DirDialog(self, "Choose output directory", "") as dirDialog:
                if dirDialog.ShowModal() == wx.ID_CANCEL:
                    return  # the user changed their mind
                pathname = dirDialog.GetPath()
                self.export_snapshot(event, multi, pathname)
            resp = wx.MessageBox('Will open directory, continue?', 'Open file', wx.OK | wx.CANCEL)
            if resp == wx.OK:
                subprocess.Popen(["explorer", os.path.abspath(pathname)])
        else:
            self.export_snapshot(event, multi, pathname)

    def onTakeSnapShot(self, event, path, page):
        """ Takes a screenshot of the screen at give pos & size (rect). """
        print('Taking screenshot...')
        rect = self.GetRect()
        # see http://aspn.activestate.com/ASPN/Mail/Message/wxpython-users/3575899
        # created by Andrea Gavana

        # adjust widths for Linux (figured out by John Torres
        # http://article.gmane.org/gmane.comp.python.wxpython/67327)
        # if sys.platform == 'linux2':
        #     client_x, client_y = self.ClientToScreen((0, 0))
        #     border_width = client_x - rect.x
        #     title_bar_height = client_y - rect.y
        #     rect.width += (border_width * 2)
        #     rect.height += title_bar_height + border_width

        # Create a DC for the whole screen area

        dcScreen = wx.ScreenDC()

        # Create a Bitmap that will hold the screenshot image later on
        # Note that the Bitmap must have a size big enough to hold the screenshot
        # -1 means using the current default colour depth
        bmp = wx.Bitmap(rect.width, rect.height)

        # Create a memory DC that will be used for actually taking the screenshot
        memDC = wx.MemoryDC()

        # Tell the memory DC to use our Bitmap
        # all drawing action on the memory DC will go to the Bitmap now
        memDC.SelectObject(bmp)

        # Blit (in this case copy) the actual screen on the memory DC
        # and thus the Bitmap
        memDC.Blit(0,  # Copy to this X coordinate
                   0,  # Copy to this Y coordinate
                   rect.width,  # Copy this width
                   rect.height,  # Copy this height
                   dcScreen,  # From where do we copy?
                   rect.x,  # What's the X offset in the original DC?
                   rect.y  # What's the Y offset in the original DC?
                   )

        # Select the Bitmap out of the memory DC by selecting a new
        # uninitialized Bitmap
        memDC.SelectObject(wx.NullBitmap)

        img = bmp.ConvertToImage()
        fileName = f"{path}/{page.wafer_info.tab_name}.png"
        img.SaveFile(fileName, wx.BITMAP_TYPE_PNG)
        print(f'saving {page.wafer_info.tab_name} as png!')

    def on_multi_export_excel(self, event):
        self.on_export_excel(event, multi=True)

    def on_export_excel(self, event, multi=False):

        with wx.FileDialog(self, "Save XLSX file",
                           wildcard="XLSX files (*.xlsx)|*.xlsx",
                           defaultFile='',
                           defaultDir=os.path.dirname(os.getcwd()),
                           style=wx.FD_SAVE | wx.FD_OVERWRITE_PROMPT) as fileDialog:

            if fileDialog.ShowModal() == wx.ID_CANCEL:
                return  # the user changed their mind

            # save the current contents in the file
            pathname = fileDialog.GetPath()
            try:
                if multi:
                    df_data_list = []
                    for i, page in enumerate(self.notebook):
                        pathname, df_data = self.single_export(page, pathname, first=False if i else True)
                        df_data_list.append(df_data)
                    else:
                        pathname = wm_utils.data_export_to_excel(df_data_list, pathname)
                        pathname.save()
                else:
                    page = self.notebook.CurrentPage
                    pathname, df_data = self.single_export(page, pathname, first=True)
                    pathname = wm_utils.data_export_to_excel([df_data], pathname)
                    pathname.save()
                # if self.notebook.CurrentPage.data_type == wm_const.DataType.DISCRETE:
                #     color_func = self.notebook.CurrentPage.legend.color_dict.get
                # else:
                #     color_func = self.notebook.CurrentPage.legend.get_color
                # die_size = self.notebook.CurrentPage.die_size
                # die_size = tuple([x / die_size[0] * 5 for x in die_size])
                # df_map = self.notebook.CurrentPage.df_map
                # wm_utils.export_to_excel(df_map, color_func, pathname, die_size, False,
                #                          self.notebook.CurrentPage.tab_name,
                #                          self.notebook.CurrentPage.plot_range)
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
        self.onExit(event)
        # self.Close(True)

    # TODO: I don't think I need a separate method for this
    def on_zoom_fit(self, event):
        """Call :meth:`wafer_map.wm_core.WaferMapPanel.zoom_fill()`."""
        print("Frame Event!")
        self.notebook.CurrentPage.zoom_fill()

    def on_next_tab(self, event):
        number_of_pages = self.notebook.GetPageCount()
        page = self.notebook.GetSelection() + 1
        if page < number_of_pages:
            self.notebook.ChangeSelection(page)

    def on_previous_tab(self, event):
        number_of_pages = self.notebook.GetPageCount()
        page = self.notebook.GetSelection() - 1
        if 0 <= page < number_of_pages:
            self.notebook.ChangeSelection(page)

    # TODO: I don't think I need a separate method for this
    def on_toggle_crosshairs(self, event):
        """Call :meth:`wafer_map.wm_core.WaferMapPanel.toggle_crosshairs()`."""
        self.notebook.CurrentPage.toggle_crosshairs()

    # TODO: I don't think I need a separate method for this
    def on_toggle_diecenters(self, event):
        """Call :meth:`wafer_map.wm_core.WaferMapPanel.toggle_crosshairs()`."""
        self.notebook.CurrentPage.toggle_die_centers()

    # TODO: I don't think I need a separate method for this
    def on_toggle_outline(self, event):
        """Call the WaferMapPanel.toggle_outline() method."""
        self.notebook.CurrentPage.toggle_outline()

    # TODO: I don't think I need a separate method for this
    #       However if I don't use these then I have to
    #           1) instance self.panel at the start of __init__
    #           2) make it so that self.panel.toggle_legend accepts arg: event
    def on_toggle_legend(self, event):
        """Call the WaferMapPanel.toggle_legend() method."""
        self.notebook.CurrentPage.toggle_legend()

    # TODO: See the 'and' in the docstring? Means I need a separate method!
    def on_change_high_color(self, event):
        """Change the high color and refresh display."""
        print("High color menu item clicked!")
        cd = wx.ColourDialog(self)
        cd.GetColourData().SetChooseFull(True)

        if cd.ShowModal() == wx.ID_OK:
            new_color = cd.GetColourData().Colour
            print("The color {} was chosen!".format(new_color))
            self.notebook.CurrentPage.on_color_change({'high': new_color, 'low': None})
            self.notebook.CurrentPage.Refresh()
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
            self.notebook.CurrentPage.on_color_change({'high': None, 'low': new_color})
            self.notebook.CurrentPage.Refresh()
        else:
            print("no color chosen :-(")
        cd.Destroy()

    def on_change_delay_time(self, event):
        dlg = wx.TextEntryDialog(self, 'Delay time (s):', 'Set capture image delay time')
        dlg.SetValue(f"{self.capture_delay_time}")
        if dlg.ShowModal() == wx.ID_OK:
            print('You entered: %s\n' % dlg.GetValue())
            self.capture_delay_time = int(dlg.GetValue())
        dlg.Destroy()


def main():
    """Run when called as a module."""
    app = wx.App()
    frame = WaferMapWindow("Testing", [], None)
    frame.Show()
    app.MainLoop()


if __name__ == "__main__":
    main()
