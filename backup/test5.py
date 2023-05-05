
import wx



""" Copied for m http://wiki.wxpython.org/Simple%20wx.Notebook%20Example """





class PageOne(wx.Panel):

    def __init__(self, parent):

        wx.Panel.__init__(self, parent)

        wx.StaticText(self, -1, "This is a PageOne object", (20,20))





class PageTwo(wx.Panel):

    def __init__(self, parent):

        wx.Panel.__init__(self, parent)

        wx.StaticText(self, -1, "This is a PageTwo object", (40, 40))





class PageThree(wx.Panel):

    def __init__(self, parent):

        wx.Panel.__init__(self, parent)

        wx.StaticText(self, -1, "This is a PageThree object", (60, 60))





class PageDynamic(wx.Panel):

    def __init__(self, parent):

        wx.Panel.__init__(self, parent)

        wx.StaticText(self, -1, "This is a Dynamic object", (60, 60))





class MainFrame(wx.Frame):

    def __init__(self):

        wx.Frame.__init__(self, None, title="Simple Notebook Example")



        # Here we create a panel and a notebook on the panel

        p = wx.Panel(self)

        self.nb = wx.Notebook(p)



        # create the page windows as children of the notebook

        page1 = PageOne(self.nb)

        page2 = PageTwo(self.nb)

        page3 = PageThree(self.nb)



        # add the pages to the notebook with the label to show on the tab

        self.nb.AddPage(page1, "Page 1")

        self.nb.AddPage(page2, "Page 2")

        self.nb.AddPage(page3, "Page 3")



        # finally, put the notebook in a sizer for the panel to manage

        # the layout

        sizer = wx.BoxSizer()

        sizer.Add(self.nb, 1, wx.EXPAND)

        p.SetSizer(sizer)



        page3.Bind(wx.EVT_LEFT_DCLICK, self.dynamic_tab)



    def dynamic_tab(self, event):

        print('dynamic_tab()')

        dynamic_page = PageDynamic(self.nb)

        self.nb.AddPage(dynamic_page, "Page Dynamic")



if __name__ == "__main__":

    app = wx.App()

    MainFrame().Show()

    app.MainLoop()