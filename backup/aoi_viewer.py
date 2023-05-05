#!/usr/bin/env python
# _*_ coding:utf-8 _*_
import glob
import os
import subprocess

import wx
from pubsub import pub


########################################################################
class ViewerPanel(wx.Panel):
    """"""

    # ----------------------------------------------------------------------
    def __init__(self, parent, yx=(0, 0), df_wafer=None):
        """Constructor"""
        wx.Panel.__init__(self, parent)
        self.parent = parent
        self.defaultColor = self.GetBackgroundColour()
        self.img = wx.Image()
        self.img_name, self.wafer_id, self.side, self.extension, self.folder, self.bin_code = ('' for t in range(6))
        self.df_wafer = df_wafer
        self.back_up_picPaths = []
        self.picPaths = []
        self.currentPicture = 0
        self.totalPictures = 0
        width, height = wx.DisplaySize()
        self.photoMaxSize = height - 200
        self.currentY, self.currentX = yx
        pub.subscribe(self.updateImages, ("update images"))
        pub.subscribe(self.updateFolder, ("update folder"))
        # pub.subscribe(self.resizeImageToFit, ('resize panel'))
        self.layout()

    # ----------------------------------------------------------------------
    # def set_photo_max_size(self):
    #     # pass
    #     width, height = wx.DisplaySize()
    #     self.photoMaxSize = height - 200

    # ----------------------------------------------------------------------
    def layout(self):
        """
        Layout the widgets on the panel
        """

        self.mainSizer = wx.BoxSizer(wx.VERTICAL)
        btnSizer = wx.BoxSizer(wx.HORIZONTAL)
        positionSizer = wx.BoxSizer(wx.HORIZONTAL)
        binSizer = wx.BoxSizer(wx.HORIZONTAL)

        img = wx.Image(self.photoMaxSize, self.photoMaxSize)
        self.imageCtrl = wx.StaticBitmap(self, wx.ID_ANY, wx.Bitmap(img))
        self.mainSizer.Add(self.imageCtrl, 0, wx.ALL | wx.CENTER, 5)
        self.imageLabel = wx.StaticText(self, label="")
        self.mainSizer.Add(self.imageLabel, 0, wx.ALL | wx.CENTER, 5)
        self.Ystr = wx.StaticText(self, label="Y:")
        self.Xstr = wx.StaticText(self, label="X:")
        self.editY = wx.TextCtrl(self, value=str(self.currentY))
        self.editX = wx.TextCtrl(self, value=str(self.currentX))
        self.editBinHint = wx.StaticText(self, label='Search specific bin image:')
        self.currentBin = wx.StaticText(self, label='Current bin:')
        self.editBin = wx.TextCtrl(self, value=str(self.bin_code))

        for text in [self.Ystr, self.editY, self.Xstr, self.editX]:
            positionSizer.Add(text, 0, wx.ALL | wx.CENTER, 5)

        for text in [self.editBinHint, self.editBin, self.currentBin]:
            binSizer.Add(text, 0, wx.ALL | wx.CENTER, 5)

        btnData = [("Previous", btnSizer, self.onPrevious),
                   ("Next", btnSizer, self.onNext),
                   ("Up", btnSizer, self.onUp),
                   ("Down", btnSizer, self.onDown),
                   ("Open file", btnSizer, self.openfile),
                   ('Search', positionSizer, self.onSearch),
                   ('Search', binSizer, self.search_specific_bin),
                   ('reset', binSizer, self.reset_to_default_paths)]

        for data in btnData:
            label, sizer, handler = data
            self.btnBuilder(label, sizer, handler)
        self.mainSizer.Add(positionSizer, 0, wx.CENTER)
        self.mainSizer.Add(binSizer, 0, wx.CENTER)
        self.mainSizer.Add(btnSizer, 0, wx.CENTER)
        self.SetSizer(self.mainSizer)

    # ----------------------------------------------------------------------
    def btnBuilder(self, label, sizer, handler):
        """
        Builds a button, binds it to an event handler and adds it to a sizer
        """
        btn = wx.Button(self, label=label)
        btn.Bind(wx.EVT_BUTTON, handler)
        sizer.Add(btn, 0, wx.ALL | wx.CENTER, 5)

    # ----------------------------------------------------------------------
    def searchImage(self, Y, X, return_img_name=False):
        # Y = self.editY.GetValue()
        # X = self.editX.GetValue()
        strY = f"{Y:.0f}".zfill(3)
        strX = f"{X:.0f}".zfill(3)
        img_name = f"{self.folder}\\{self.wafer_id}_{self.side}_{strY}_{strX}.{self.extension}"
        if return_img_name:
            return img_name
        try:
            self.currentPicture = self.picPaths.index(img_name)
            self.loadImage(img_name)
        except ValueError:
            self.loadImage(self.picPaths[0])
            msg = f'Cant find {self.wafer_id},\n {img_name}, please reconfirm.\n' \
                f'Only find {self.totalPictures} images in the folder.'
            resp = wx.MessageBox(msg, 'Warning', wx.OK)
            print('No image found')

    # ----------------------------------------------------------------------
    def search_specific_bin(self, event):
        bin_code = self.editBin.GetValue()
        df = self.df_wafer.loc[self.df_wafer['DATA'] == bin_code]
        img_list = []
        for X, Y, DATA in df.itertuples(index=False):
            img_list.append(self.searchImage(Y, X, return_img_name=True))
        self.back_up_picPaths = self.picPaths
        self.picPaths = img_list
        self.picPaths.sort()
        self.totalPictures = len(self.picPaths)
        self.currentPicture = 0
        img_name = self.picPaths[0]
        self.extension = img_name.split('.')[-1]
        self.loadImage(img_name)

    # ----------------------------------------------------------------------
    def reset_to_default_paths(self, event):
        if self.back_up_picPaths:
            self.picPaths = self.back_up_picPaths

        self.updateImages(msg='')
        self.editBin.SetValue('')

    def scaleImage(self):
        # scale the image, preserving the aspect ratio
        width, height = wx.DisplaySize()
        # width, height = self.parent.GetSize()
        if width > height:
            self.photoMaxSize = height
        else:
            self.photoMaxSize = width
        W = self.img.GetWidth()
        H = self.img.GetHeight()
        if W > H:
            NewW = self.photoMaxSize
            NewH = self.photoMaxSize * H / W
        else:
            NewH = self.photoMaxSize
            NewW = self.photoMaxSize * W / H
        self.img = self.img.Scale(NewW, NewH)

        self.imageCtrl.SetBitmap(wx.Bitmap(self.img))
        pub.sendMessage("resize", msg='')
        self.Refresh()

    # ----------------------------------------------------------------------
    def loadImage(self, image):
        """"""
        image_name = os.path.basename(image)
        df = self.df_wafer
        name, extension = image_name.split('.')
        Y, X = name.split('_')[-2:]
        self.bin_code = df[(df['Y'] == int(Y)) & (df['X'] == int(X))]['DATA'].values[0]

        self.currentBin.SetLabel(f'Current bin:{self.bin_code}')
        self.imageLabel.SetLabel(image_name)
        self.img_name = image_name
        self.img = wx.Image(image, wx.BITMAP_TYPE_ANY)
        # scale the image, preserving the aspect ratio
        self.scaleImage()
        # W = self.img.GetWidth()
        # H = self.img.GetHeight()
        # if W > H:
        #     NewW = self.photoMaxSize
        #     NewH = self.photoMaxSize * H / W
        # else:
        #     NewH = self.photoMaxSize
        #     NewW = self.photoMaxSize * W / H
        # self.img = self.img.Scale(NewW, NewH)
        #
        # self.imageCtrl.SetBitmap(wx.Bitmap(self.img))
        # pub.sendMessage("resize", msg='')
        # self.Refresh()

    # ----------------------------------------------------------------------
    def getCurrentPosition(self):
        name, extension = self.img_name.split('.')
        Y, X = name.split('_')[-2:]
        self.currentY = int(Y)
        self.currentX = int(X)

    def upPicture(self):
        self.getCurrentPosition()
        Y = self.currentY + 1
        self.searchImage(Y, self.currentX)

    def downPicture(self):
        self.getCurrentPosition()
        Y = self.currentY - 1
        self.searchImage(Y, self.currentX)

    def nextPicture(self):
        """
        Loads the next picture in the directory
        """
        if self.currentPicture == self.totalPictures - 1:
            self.currentPicture = 0
        else:
            self.currentPicture += 1
        self.loadImage(self.picPaths[self.currentPicture])

    # ----------------------------------------------------------------------
    def previousPicture(self):
        """
        Displays the previous picture in the directory
        """
        if self.currentPicture == 0:
            self.currentPicture = self.totalPictures - 1
        else:
            self.currentPicture -= 1
        self.loadImage(self.picPaths[self.currentPicture])

    # ----------------------------------------------------------------------
    # def update(self, event):
    #     """
    #     Called when the slideTimer's timer event fires. Loads the next
    #     picture from the folder by calling th nextPicture method
    #     """
    #     self.nextPicture()

    # ----------------------------------------------------------------------
    def updateFolder(self, msg):
        self.folder, self.picPaths = msg
        nlst = self.folder.split('_')
        name = nlst[0]
        # side = '_'.join(nlst[1:])
        wafer_id = name.split('\\')[-1:][0]
        self.wafer_id = wafer_id
        filename = self.picPaths[0]
        self.side = '_'.join([s for s in filename.split('\\')[-1].split('_') if s.isalpha()])

    def updateImages(self, msg):
        """
        Updates the picPaths list to contain the current folder's images
        """
        # self.picPaths = msg
        self.picPaths.sort()
        self.totalPictures = len(self.picPaths)
        img_name = self.picPaths[0]
        self.extension = img_name.split('.')[-1]
        if self.currentX and self.currentY:
            self.searchImage(self.currentY, self.currentX)
        else:
            self.loadImage(img_name)

    # def on_quit(self, msg):
    #     self.Close()

    # ----------------------------------------------------------------------
    def onNext(self, event):
        """
        Calls the nextPicture method
        """
        self.nextPicture()

    # ----------------------------------------------------------------------
    def onPrevious(self, event):
        """
        Calls the previousPicture method
        """
        self.previousPicture()

    def onUp(self, event):
        self.upPicture()

    def onDown(self, event):
        self.downPicture()

    def onSearch(self, event):
        self.searchImage(int(self.editY.GetValue()), int(self.editX.GetValue()))

    def openfile(self, event):
        subprocess.Popen([os.path.abspath(f'{self.folder}\\{self.img_name}')], shell=True)


class AOIViewerFrame(wx.Frame):
    """"""

    # ----------------------------------------------------------------------
    def __init__(self, default_path='', yx=(0, 0), pic_paths=None, parent=None, df_wafer=None):
        """Constructor"""
        wx.Frame.__init__(self, parent, title="AOI pictures Viewer"
                          ,
                          style=wx.DEFAULT_FRAME_STYLE ^ wx.RESIZE_BORDER ^ wx.MAXIMIZE_BOX)
        # self.panel = ViewerPanel(self, yx=yx)
        self.Maximize()
        self.Bind(wx.EVT_CLOSE, self.on_quit)
        self.folderPath = default_path
        pub.subscribe(self.resizeFrame, ("resize"))
        if pic_paths is None:
            self.picPaths = []
        else:
            self.picPaths = pic_paths
        self.initToolbar()
        self.sizer = wx.BoxSizer(wx.VERTICAL)
        # init the panel setting
        self.panel = ViewerPanel(self, yx=yx, df_wafer=df_wafer)
        self.sizer.Add(self.panel, 1, wx.EXPAND)
        self.SetSizer(self.sizer)
        self.Show()
        self.sizer.Fit(self)
        self.Center()
        self.OpenDirectory()
        self.Bind(wx.EVT_SIZE, self.resizeFrame)

    # ----------------------------------------------------------------------
    def initToolbar(self):
        """
        Initialize the toolbar
        """
        self.toolbar = self.CreateToolBar()
        self.toolbar.SetToolBitmapSize((16, 16))

        open_ico = wx.ArtProvider.GetBitmap(wx.ART_FILE_OPEN, wx.ART_TOOLBAR, (16, 16))
        openTool = self.toolbar.AddTool(toolId=wx.ID_ANY, bitmap=open_ico, label="Open",
                                        shortHelp="Open an Image Directory")
        self.Bind(wx.EVT_MENU, self.onOpenDirectory, openTool)

        self.toolbar.Realize()

    # ----------------------------------------------------------------------
    # def onResize(self, event):

    #     # width, height = wx.DisplaySize()
    #     width, height = self.GetSize()
    #     photoMaxSize = height - 10
    #     W = self.panel.img.GetWidth()
    #     H = self.panel.img.GetHeight()
    #     if W > H:
    #         NewW = photoMaxSize
    #         NewH = photoMaxSize * H / W
    #     else:
    #         NewH = photoMaxSize
    #         NewW = photoMaxSize * W / H
    #     # frame_size = self.GetSize()
    #     # frame_h = (frame_size[0] - 10) / 2
    #     # frame_w = (frame_size[1] - 10) / 2
    #     img = self.panel.img.Scale(NewW, NewH)
    #     self.panel.imageCtrl.SetBitmap(wx.Bitmap(img))
    #     self.panel.Refresh()
    #     # self.Layout()

    # ----------------------------------------------------------------------
    def onOpenDirectory(self, event):
        """
        Opens a DirDialog to allow the user to open a folder with pictures
        """
        dlg = wx.DirDialog(self, "Choose a directory",
                           style=wx.DD_DEFAULT_STYLE)

        if dlg.ShowModal() == wx.ID_OK:
            self.folderPath = dlg.GetPath()
            self.OpenDirectory()

    def OpenDirectory(self):
        print(self.folderPath)
        if self.picPaths:
            picPaths = self.picPaths
        else:
            picPaths = glob.glob(self.folderPath + "\\*.jpg")
        print(picPaths)
        pub.sendMessage("update folder", msg=(self.folderPath, picPaths))
        pub.sendMessage("update images", msg='')

    def on_quit(self, event):
        """Action for the quit event."""

        pub.unsubscribe(self.resizeFrame, ("resize"))
        pub.unsubAll()
        event.Skip()

    # ----------------------------------------------------------------------
    def resizeFrame(self, msg):
        """"""
        # pass
        # self.panel.scaleImage()
        self.sizer.Fit(self)


# ----------------------------------------------------------------------
class AOIViewer(object):

    def __init__(self, default_path='', yx=(0, 0), main_app=False, pic_paths=None, parent=None, df_wafer=None):
        if main_app:
            app = wx.App()
        self.default_path = default_path
        self.yx = yx
        self.pic_paths = pic_paths
        self.frame = AOIViewerFrame(self.default_path, self.yx, pic_paths=self.pic_paths, parent=parent,
                                    df_wafer=df_wafer)

        self.frame.SetIcon(wx.Icon('icon3.png'))
        self.frame.Show()
        if main_app:
            app.MainLoop()


def main():
    app = wx.App()
    frame = AOIViewerFrame(default_path='\\\\192.168.7.27\\\\aoi3\\1NJK967#14_FrontSide', yx=(1, 24))
    app.MainLoop()


# ----------------------------------------------------------------------
if __name__ == "__main__":
    main()
