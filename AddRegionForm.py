# 1.0
import json

import System.Drawing
import System.Windows.Forms

from System.Drawing import *
from System.Windows.Forms import *

import MapHelper

class AddRegionForm(Form):
    def __init__(self, parent, editing=False):
        self._parent = parent
        self.InitializeComponent()

        self._editing = False
        if editing != False:
            self._editing = True
            self._editRegion = editing

        self._mwIncrement = 0.02

        self.InitializeData()

    def InitializeComponent(self):
        self._wbEditor = System.Windows.Forms.WebBrowser()
        self._wbViewer = System.Windows.Forms.WebBrowser()
        self._cbRegions = System.Windows.Forms.ComboBox()
        self._lvPoints = System.Windows.Forms.ListView()
        self._tbFColor = System.Windows.Forms.TextBox()
        self._lblFColor = System.Windows.Forms.Label()
        self._colorPicker = System.Windows.Forms.ColorDialog()
        self._lblFOpacity = System.Windows.Forms.Label()
        self._tbFOpacity = System.Windows.Forms.TextBox()
        self._lblSColor = System.Windows.Forms.Label()
        self._tbSColor = System.Windows.Forms.TextBox()
        self._lblSOpacity = System.Windows.Forms.Label()
        self._tbSOpacity = System.Windows.Forms.TextBox()
        self._lblRegionName = System.Windows.Forms.Label()
        self._tbRegionName = System.Windows.Forms.TextBox()
        self._btnSave = System.Windows.Forms.Button()
        self.SuspendLayout()
        #
        # wbEditor
        #
        self._wbEditor.AllowWebBrowserDrop = False
        self._wbEditor.Anchor = System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left | System.Windows.Forms.AnchorStyles.Right
        self._wbEditor.IsWebBrowserContextMenuEnabled = False
        self._wbEditor.Location = System.Drawing.Point(0, 0)
        self._wbEditor.MinimumSize = System.Drawing.Size(20, 20)
        self._wbEditor.Name = "wbEditor"
        self._wbEditor.ScrollBarsEnabled = False
        self._wbEditor.Size = System.Drawing.Size(473, 612)
        self._wbEditor.TabIndex = 0
        #
        # wbViewer
        #
        self._wbViewer.Dock = System.Windows.Forms.DockStyle.Right
        self._wbViewer.Location = System.Drawing.Point(697, 0)
        self._wbViewer.MinimumSize = System.Drawing.Size(20, 20)
        self._wbViewer.Name = "wbViewer"
        self._wbViewer.Size = System.Drawing.Size(287, 612)
        self._wbViewer.TabIndex = 1
        #
        # cbRegions
        #
        self._cbRegions.Anchor = System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right
        self._cbRegions.FormattingEnabled = True
        self._cbRegions.Location = System.Drawing.Point(479, 3)
        self._cbRegions.Name = "cbRegions"
        self._cbRegions.Size = System.Drawing.Size(212, 21)
        self._cbRegions.TabIndex = 2
        self._cbRegions.SelectedIndexChanged += self.CbRegionsSelectedIndexChanged
        #
        # lvPoints
        #
        self._lvPoints.Anchor = System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right
        self._lvPoints.FullRowSelect = True
        self._lvPoints.GridLines = True
        self._lvPoints.Location = System.Drawing.Point(479, 30)
        self._lvPoints.Name = "lvPoints"
        self._lvPoints.Size = System.Drawing.Size(212, 424)
        self._lvPoints.TabIndex = 3
        self._lvPoints.UseCompatibleStateImageBehavior = False
        self._lvPoints.View = System.Windows.Forms.View.Details
        self._lvPoints.MouseClick += self.LvPointsMouseClick
        #
        # tbFColor
        #
        self._tbFColor.Anchor = System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right
        self._tbFColor.Location = System.Drawing.Point(567, 486)
        self._tbFColor.Name = "tbFColor"
        self._tbFColor.ReadOnly = True
        self._tbFColor.Size = System.Drawing.Size(124, 20)
        self._tbFColor.TabIndex = 4
        self._tbFColor.MouseClick += self.TbColorMouseClick
        #
        # lblFColor
        #
        self._lblFColor.Anchor = System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right
        self._lblFColor.Location = System.Drawing.Point(479, 483)
        self._lblFColor.Name = "lblFColor"
        self._lblFColor.Size = System.Drawing.Size(82, 23)
        self._lblFColor.TabIndex = 5
        self._lblFColor.Text = "Fill Color:"
        self._lblFColor.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        #
        # lblFOpacity
        #
        self._lblFOpacity.Anchor = System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right
        self._lblFOpacity.Location = System.Drawing.Point(479, 509)
        self._lblFOpacity.Name = "lblFOpacity"
        self._lblFOpacity.Size = System.Drawing.Size(82, 23)
        self._lblFOpacity.TabIndex = 7
        self._lblFOpacity.Text = "Fill Opacity:"
        self._lblFOpacity.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        #
        # tbFOpacity
        #
        self._tbFOpacity.Anchor = System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right
        self._tbFOpacity.Location = System.Drawing.Point(567, 512)
        self._tbFOpacity.Name = "tbFOpacity"
        self._tbFOpacity.ReadOnly = True
        self._tbFOpacity.Size = System.Drawing.Size(124, 20)
        self._tbFOpacity.TabIndex = 6
        self._tbFOpacity.MouseWheel += self.TbOpacityMouseWheel
        #
        # lblSColor
        #
        self._lblSColor.Anchor = System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right
        self._lblSColor.Location = System.Drawing.Point(479, 535)
        self._lblSColor.Name = "lblSColor"
        self._lblSColor.Size = System.Drawing.Size(82, 23)
        self._lblSColor.TabIndex = 9
        self._lblSColor.Text = "Stroke Color:"
        self._lblSColor.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        #
        # tbSColor
        #
        self._tbSColor.Anchor = System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right
        self._tbSColor.Location = System.Drawing.Point(567, 538)
        self._tbSColor.Name = "tbSColor"
        self._tbSColor.ReadOnly = True
        self._tbSColor.Size = System.Drawing.Size(124, 20)
        self._tbSColor.TabIndex = 8
        self._tbSColor.MouseClick += self.TbColorMouseClick
        #
        # lblSOpacity
        #
        self._lblSOpacity.Anchor = System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right
        self._lblSOpacity.Location = System.Drawing.Point(479, 561)
        self._lblSOpacity.Name = "lblSOpacity"
        self._lblSOpacity.Size = System.Drawing.Size(82, 23)
        self._lblSOpacity.TabIndex = 11
        self._lblSOpacity.Text = "Stroke Opacity:"
        self._lblSOpacity.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        #
        # tbSOpacity
        #
        self._tbSOpacity.Anchor = System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right
        self._tbSOpacity.Location = System.Drawing.Point(567, 564)
        self._tbSOpacity.Name = "tbSOpacity"
        self._tbSOpacity.ReadOnly = True
        self._tbSOpacity.Size = System.Drawing.Size(124, 20)
        self._tbSOpacity.TabIndex = 10
        self._tbSOpacity.MouseWheel += self.TbOpacityMouseWheel
        #
        # lblRegionName
        #
        self._lblRegionName.Anchor = System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right
        self._lblRegionName.Location = System.Drawing.Point(479, 457)
        self._lblRegionName.Name = "lblRegionName"
        self._lblRegionName.Size = System.Drawing.Size(82, 23)
        self._lblRegionName.TabIndex = 13
        self._lblRegionName.Text = "Name:"
        self._lblRegionName.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        #
        # tbRegionName
        #
        self._tbRegionName.Anchor = System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right
        self._tbRegionName.Location = System.Drawing.Point(567, 460)
        self._tbRegionName.Name = "tbRegionName"
        self._tbRegionName.Size = System.Drawing.Size(124, 20)
        self._tbRegionName.TabIndex = 12
        #
        # btnSave
        #
        self._btnSave.Anchor = System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right
        self._btnSave.Location = System.Drawing.Point(479, 587)
        self._btnSave.Name = "btnSave"
        self._btnSave.Size = System.Drawing.Size(212, 23)
        self._btnSave.TabIndex = 14
        self._btnSave.Text = "Save"
        self._btnSave.UseVisualStyleBackColor = True
        self._btnSave.Click += self.BtnSaveClick
        #
        # AddRegionForm
        #
        self.ClientSize = System.Drawing.Size(984, 612)
        self.Controls.Add(self._btnSave)
        self.Controls.Add(self._lblRegionName)
        self.Controls.Add(self._tbRegionName)
        self.Controls.Add(self._lblSOpacity)
        self.Controls.Add(self._tbSOpacity)
        self.Controls.Add(self._lblSColor)
        self.Controls.Add(self._tbSColor)
        self.Controls.Add(self._lblFOpacity)
        self.Controls.Add(self._tbFOpacity)
        self.Controls.Add(self._lblFColor)
        self.Controls.Add(self._tbFColor)
        self.Controls.Add(self._lvPoints)
        self.Controls.Add(self._cbRegions)
        self.Controls.Add(self._wbViewer)
        self.Controls.Add(self._wbEditor)
        self.Name = "AddRegionForm"
        self.Text = "Add Region"
        self.FormClosing += self.AddRegionFormFormClosing
        self.ResumeLayout(False)
        self.PerformLayout()

    def InitializeData( self ):
        self._lvPoints.Columns.Add('', -1)
        self._lvPoints.Columns.Add('Latitude', -1)
        self._lvPoints.Columns.Add('Longitude', -1)

        self.editorMapHelper = MapHelper.MapHelper( 'RegionEditor', self._wbEditor, MapHelper.MAP_TYPE_REGION_HELPER )
        self.viewerMapHelper = MapHelper.MapHelper( 'RegopmViewer', self._wbViewer, MapHelper.MAP_TYPE_NORMAL )

        regions = self._parent.dbHelper.getAreas()

        if regions:
            for region in regions:
                self._cbRegions.Items.Add( region[1] )

            self._cbRegions.SelectedIndex = 0

            self.viewerMapHelper.buildPage( regions[0], self.viewerMapHelper.getRegionPointsAsPeople( regions[0] ) )

        else:
            self.viewerMapHelper.buildPage()


        if self._editing:
            try:
                # print self._editRegion
                region = self._parent.dbHelper.getArea( self._editRegion )

                # print region

                self._tbRegionName.ReadOnly = True
                self._tbRegionName.Text = region[1]
                self._tbSColor.Text     = region[2]
                self._tbSOpacity.Text   = str(region[3])
                self._tbFColor.Text     = region[4]
                self._tbFOpacity.Text   = str(region[5])

                self.editorMapHelper.buildPage( region )
            except Exception, e:
                print e
                raise e

        else:
            if regions:
                points = json.loads( regions[0][6] )
                self.editorMapHelper.buildPage( clat=points[0][0], clng=points[0][1] )
            else:
                self.editorMapHelper.buildPage()
            self._tbFColor.Text   = '#999999'
            self._tbFOpacity.Text = '0.1'
            self._tbSColor.Text   = '#999999'
            self._tbSOpacity.Text = '0.7'

    def UpdateViewedRegion( self, region ):
        points = json.loads( region[6] )

        self._lvPoints.Items.Clear()

        for i in xrange(len(points)):
            char = self._parent.mapHelperRegion.peopleCharacters[i]
            self._lvPoints.Items.Add( ListViewItem((char, str(points[i][0]), str(points[i][1]))) )

        self.viewerMapHelper.buildPage( region, self.viewerMapHelper.getRegionPointsAsPeople( region ) )

    def LvPointsMouseClick(self, sender, e):
        row = [item.Text for item in self._lvPoints.SelectedItems[0].SubItems]

        Clipboard.SetText( "{}, {}".format( row[1], row[2] ) )

    def CbRegionsSelectedIndexChanged(self, sender, e):
        self.UpdateViewedRegion( self._parent.dbHelper.getArea( self._cbRegions.Items[self._cbRegions.SelectedIndex] ) )

    def rgb_to_hex(self, r, g, b):
        h = '#%02x%02x%02x' % ( r, g, b )
        return h.upper()

    def TbColorMouseClick(self, sender, e):
        if self._colorPicker.ShowDialog() == DialogResult.OK:
            sender.Text = self.rgb_to_hex( self._colorPicker.Color.R, self._colorPicker.Color.G, self._colorPicker.Color.B )
            self.wbEditorInjectValues()

    def TbOpacityMouseWheel(self, sender, e):
        if e.Delta > 0:
            opacity = float(sender.Text)
            newOpacity = opacity + self._mwIncrement
            if newOpacity <= 1.00:
                sender.Text = str( newOpacity )
                self.wbEditorInjectValues()
        elif e.Delta < 0:
            opacity = float(sender.Text)
            newOpacity = opacity - self._mwIncrement
            if newOpacity >= 0.00:
                sender.Text = str( newOpacity )
                self.wbEditorInjectValues()

    def wbEditorInjectValues( self ):
        self._wbEditor.Document.InvokeScript( "updateInfo", ( self._tbFColor.Text, float(self._tbFOpacity.Text), self._tbSColor.Text, float(self._tbSOpacity.Text) ) )
        self._wbEditor.Document.InvokeScript( "drawPolygon" )

    def BtnSaveClick( self, sender, e ):
        if self._tbRegionName.Text == '' or len(self._tbRegionName.Text) < 1:
            MessageBox.Show( 'You need a region name!', 'Whoops!', MessageBoxButtons.OK, MessageBoxIcon.Error )
            return None

        if not self._editing and self._tbRegionName.Text in self._cbRegions.Items:
            MessageBox.Show( 'The region named "%s" already exists!' % self._tbRegionName.Text, 'Error!', MessageBoxButtons.OK, MessageBoxIcon.Error )
            return None

        coords = self._wbEditor.Document.GetElementById("coordinates").InnerText

        if coords == None:
            MessageBox.Show( "You havn't picked any coordinates yet!", 'Error!', MessageBoxButtons.OK, MessageBoxIcon.Error )
            return None

        coords = coords.strip().split("\n")

        if len(coords) < 3:
            MessageBox.Show( "You need at least 3 coordinates to make an area!", 'Error!', MessageBoxButtons.OK, MessageBoxIcon.Error )
            return None

        try:
            coords = [ [ float( c.strip() ) for c in coord.split(',') ] for coord in coords ]
        except ValueError, e:
            MessageBox.Show( "Your coordinates are not formatted properly!", 'Error!', MessageBoxButtons.OK, MessageBoxIcon.Error )
            return None


        if self._editing:
            area = self._parent.dbHelper.getArea( self._editRegion )
            self._parent.dbHelper.updateArea(
                area[0],
                self._tbRegionName.Text,
                self._tbSColor.Text,
                self._tbSOpacity.Text,
                self._tbFColor.Text,
                self._tbFOpacity.Text,
                coords
            )
        else:
            self._parent.dbHelper.addArea(
                self._tbRegionName.Text,
                self._tbSColor.Text,
                self._tbSOpacity.Text,
                self._tbFColor.Text,
                self._tbFOpacity.Text,
                coords
            )

        self._parent.refreshAreas()

        self._parent.dbHelper.updatePeopleAreas()

        self.Close()

    def AddRegionFormFormClosing(self, sender, e):
        self.editorMapHelper.cleanup()
        self.viewerMapHelper.cleanup()
