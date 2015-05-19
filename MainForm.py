# 1.1
import itertools
import os

import System.Uri
import System.Drawing
import System.Windows.Forms

from operator import itemgetter
from System.Drawing import *
from System.IO import Directory, Path
from System.Windows.Forms import *

import MapHelper
import DatabaseHelper
from AddPeopleForm import AddPeopleForm
from AddRegionForm import AddRegionForm

class MainForm(Form):
    def __init__(self):
        self.InitializeComponent()

        self.mapHelperMain   = MapHelper.MapHelper( 'Main',   self._wbMap,     MapHelper.MAP_TYPE_NORMAL )
        self.mapHelperRegion = MapHelper.MapHelper( 'Region', self._wbRegions, MapHelper.MAP_TYPE_NORMAL )
        self.dbHelper        = DatabaseHelper.DatabaseHelper()

        self.InitializeData()

    def InitializeComponent(self):
        self._tcMain = System.Windows.Forms.TabControl()
        self._tpPeople = System.Windows.Forms.TabPage()
        self._dgvPeople = System.Windows.Forms.DataGridView()
        self._lblSearch = System.Windows.Forms.Label()
        self._tbSearch = System.Windows.Forms.TextBox()
        self._tpMap = System.Windows.Forms.TabPage()
        self._tpRegions = System.Windows.Forms.TabPage()
        self._wbMap = System.Windows.Forms.WebBrowser()
        self._lblResults = System.Windows.Forms.Label()
        self._tbPeopleCount = System.Windows.Forms.TextBox()
        self._btnAddPeople = System.Windows.Forms.Button()
        self._lbRegions = System.Windows.Forms.ListBox()
        self._wbRegions = System.Windows.Forms.WebBrowser()
        self._btnRegionEdit = System.Windows.Forms.Button()
        self._btnRegionNew = System.Windows.Forms.Button()
        self._lbRegionsMain = System.Windows.Forms.ListBox()
        self._lbPeopleTypes = System.Windows.Forms.ListBox()
        self._ID = System.Windows.Forms.DataGridViewTextBoxColumn()
        self._Group = System.Windows.Forms.DataGridViewComboBoxColumn()
        self._Area = System.Windows.Forms.DataGridViewTextBoxColumn()
        self._clFName = System.Windows.Forms.DataGridViewTextBoxColumn()
        self._clLName = System.Windows.Forms.DataGridViewTextBoxColumn()
        self._clPhone = System.Windows.Forms.DataGridViewTextBoxColumn()
        self._clAddress = System.Windows.Forms.DataGridViewTextBoxColumn()
        self._Comments = System.Windows.Forms.DataGridViewTextBoxColumn()
        self._Latitude = System.Windows.Forms.DataGridViewTextBoxColumn()
        self._Longitude = System.Windows.Forms.DataGridViewTextBoxColumn()
        self._btnDelete = System.Windows.Forms.Button()
        self._tcMain.SuspendLayout()
        self._tpPeople.SuspendLayout()
        self._dgvPeople.BeginInit()
        self._tpMap.SuspendLayout()
        self._tpRegions.SuspendLayout()
        self.SuspendLayout()
        #
        # tcMain
        #
        self._tcMain.Controls.Add(self._tpPeople)
        self._tcMain.Controls.Add(self._tpRegions)
        self._tcMain.Controls.Add(self._tpMap)
        self._tcMain.Dock = System.Windows.Forms.DockStyle.Fill
        self._tcMain.Location = System.Drawing.Point(0, 0)
        self._tcMain.Name = "tcMain"
        self._tcMain.SelectedIndex = 0
        self._tcMain.Size = System.Drawing.Size(984, 612)
        self._tcMain.TabIndex = 0
        #
        # tpPeople
        #
        self._tpPeople.Controls.Add(self._btnAddPeople)
        self._tpPeople.Controls.Add(self._tbPeopleCount)
        self._tpPeople.Controls.Add(self._lblResults)
        self._tpPeople.Controls.Add(self._tbSearch)
        self._tpPeople.Controls.Add(self._lblSearch)
        self._tpPeople.Controls.Add(self._dgvPeople)
        self._tpPeople.Location = System.Drawing.Point(4, 22)
        self._tpPeople.Name = "tpPeople"
        self._tpPeople.Padding = System.Windows.Forms.Padding(3)
        self._tpPeople.Size = System.Drawing.Size(976, 586)
        self._tpPeople.TabIndex = 0
        self._tpPeople.Text = "People"
        self._tpPeople.UseVisualStyleBackColor = True
        #
        # dgvPeople
        #
        self._dgvPeople.AllowUserToAddRows = False
        self._dgvPeople.AllowUserToResizeRows = False
        self._dgvPeople.Anchor = System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left | System.Windows.Forms.AnchorStyles.Right
        self._dgvPeople.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        self._dgvPeople.Columns.AddRange(System.Array[System.Windows.Forms.DataGridViewColumn](
            [self._ID,
            self._Group,
            self._Area,
            self._clFName,
            self._clLName,
            self._clPhone,
            self._clAddress,
            self._Comments,
            self._Latitude,
            self._Longitude]))
        self._dgvPeople.Location = System.Drawing.Point(3, 29)
        self._dgvPeople.Name = "dgvPeople"
        self._dgvPeople.RowTemplate.Resizable = System.Windows.Forms.DataGridViewTriState.False
        self._dgvPeople.Size = System.Drawing.Size(970, 554)
        self._dgvPeople.TabIndex = 0
        self._dgvPeople.CellEndEdit += self.DgvPeopleCellEndEdit
        self._dgvPeople.CellValidating += self.DgvPeopleCellValidating
        self._dgvPeople.UserDeletedRow += self.DgvPeopleUserDeletedRow
        self._dgvPeople.UserDeletingRow += self.DgvPeopleUserDeletingRow
        self._dgvPeople.RowEnter += self.DgvPeopleRowEnter
        self._dgvPeople.RowLeave += self.DgvPeopleRowLeave
        self._dgvPeople.CellClick += self.DgvPeopleCellClick
        #
        # lblSearch
        #
        self._lblSearch.Location = System.Drawing.Point(4, 6)
        self._lblSearch.Name = "lblSearch"
        self._lblSearch.Size = System.Drawing.Size(44, 17)
        self._lblSearch.TabIndex = 1
        self._lblSearch.Text = "Search:"
        #
        # tbSearch
        #
        self._tbSearch.Anchor = System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left | System.Windows.Forms.AnchorStyles.Right
        self._tbSearch.Location = System.Drawing.Point(45, 3)
        self._tbSearch.Name = "tbSearch"
        self._tbSearch.Size = System.Drawing.Size(679, 20)
        self._tbSearch.TabIndex = 2
        self._tbSearch.KeyUp += self.TbSearchKeyUp
        #
        # tpMap
        #
        self._tpMap.Controls.Add(self._lbPeopleTypes)
        self._tpMap.Controls.Add(self._lbRegionsMain)
        self._tpMap.Controls.Add(self._wbMap)
        self._tpMap.Location = System.Drawing.Point(4, 22)
        self._tpMap.Name = "tpMap"
        self._tpMap.Padding = System.Windows.Forms.Padding(3)
        self._tpMap.Size = System.Drawing.Size(976, 586)
        self._tpMap.TabIndex = 2
        self._tpMap.Text = "Map"
        self._tpMap.UseVisualStyleBackColor = True
        #
        # tpRegions
        #
        self._tpRegions.Controls.Add(self._btnDelete)
        self._tpRegions.Controls.Add(self._btnRegionNew)
        self._tpRegions.Controls.Add(self._btnRegionEdit)
        self._tpRegions.Controls.Add(self._wbRegions)
        self._tpRegions.Controls.Add(self._lbRegions)
        self._tpRegions.Location = System.Drawing.Point(4, 22)
        self._tpRegions.Name = "tpRegions"
        self._tpRegions.Padding = System.Windows.Forms.Padding(3)
        self._tpRegions.Size = System.Drawing.Size(976, 586)
        self._tpRegions.TabIndex = 3
        self._tpRegions.Text = "Regions"
        self._tpRegions.UseVisualStyleBackColor = True
        #
        # wbMap
        #
        self._wbMap.AllowWebBrowserDrop = False
        self._wbMap.Anchor = System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left | System.Windows.Forms.AnchorStyles.Right
        self._wbMap.IsWebBrowserContextMenuEnabled = False
        self._wbMap.Location = System.Drawing.Point(204, 3)
        self._wbMap.MinimumSize = System.Drawing.Size(20, 20)
        self._wbMap.Name = "wbMap"
        self._wbMap.ScriptErrorsSuppressed = True
        self._wbMap.ScrollBarsEnabled = False
        self._wbMap.Size = System.Drawing.Size(772, 580)
        self._wbMap.TabIndex = 0
        self._wbMap.Url = System.Uri("", System.UriKind.Relative)
        #
        # lblResults
        #
        self._lblResults.Anchor = System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right
        self._lblResults.Location = System.Drawing.Point(730, 6)
        self._lblResults.Name = "lblResults"
        self._lblResults.Size = System.Drawing.Size(38, 20)
        self._lblResults.TabIndex = 3
        self._lblResults.Text = "Count:"
        #
        # tbPeopleCount
        #
        self._tbPeopleCount.Anchor = System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right
        self._tbPeopleCount.Location = System.Drawing.Point(766, 3)
        self._tbPeopleCount.Name = "tbPeopleCount"
        self._tbPeopleCount.ReadOnly = True
        self._tbPeopleCount.Size = System.Drawing.Size(121, 20)
        self._tbPeopleCount.TabIndex = 4
        #
        # btnAddPeople
        #
        self._btnAddPeople.Anchor = System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right
        self._btnAddPeople.Location = System.Drawing.Point(893, 3)
        self._btnAddPeople.Name = "btnAddPeople"
        self._btnAddPeople.Size = System.Drawing.Size(75, 23)
        self._btnAddPeople.TabIndex = 5
        self._btnAddPeople.Text = "AddPeople"
        self._btnAddPeople.UseVisualStyleBackColor = True
        self._btnAddPeople.Click += self.BtnAddPeopleClick
        #
        # lbRegions
        #
        self._lbRegions.Anchor = System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left
        self._lbRegions.FormattingEnabled = True
        self._lbRegions.Location = System.Drawing.Point(6, 6)
        self._lbRegions.Name = "lbRegions"
        self._lbRegions.Size = System.Drawing.Size(192, 550)
        self._lbRegions.TabIndex = 0
        self._lbRegions.SelectedIndexChanged += self.LbRegionsSelectedIndexChanged
        #
        # wbRegions
        #
        self._wbRegions.AllowWebBrowserDrop = False
        self._wbRegions.Anchor = System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left | System.Windows.Forms.AnchorStyles.Right
        self._wbRegions.IsWebBrowserContextMenuEnabled = False
        self._wbRegions.Location = System.Drawing.Point(204, 6)
        self._wbRegions.MinimumSize = System.Drawing.Size(20, 20)
        self._wbRegions.Name = "wbRegions"
        self._wbRegions.ScrollBarsEnabled = False
        self._wbRegions.Size = System.Drawing.Size(769, 574)
        self._wbRegions.TabIndex = 1
        #
        # btnRegionEdit
        #
        self._btnRegionEdit.Anchor = System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left
        self._btnRegionEdit.Location = System.Drawing.Point(68, 560)
        self._btnRegionEdit.Name = "btnRegionEdit"
        self._btnRegionEdit.Size = System.Drawing.Size(62, 23)
        self._btnRegionEdit.TabIndex = 2
        self._btnRegionEdit.Text = "Edit"
        self._btnRegionEdit.UseVisualStyleBackColor = True
        self._btnRegionEdit.Click += self.BtnRegionEditClick
        #
        # btnRegionNew
        #
        self._btnRegionNew.Anchor = System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left
        self._btnRegionNew.Location = System.Drawing.Point(136, 560)
        self._btnRegionNew.Name = "btnRegionNew"
        self._btnRegionNew.Size = System.Drawing.Size(62, 23)
        self._btnRegionNew.TabIndex = 3
        self._btnRegionNew.Text = "New"
        self._btnRegionNew.UseVisualStyleBackColor = True
        self._btnRegionNew.Click += self.BtnRegionNewClick
        #
        # lbRegionsMain
        #
        self._lbRegionsMain.Anchor = System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left
        self._lbRegionsMain.FormattingEnabled = True
        self._lbRegionsMain.Location = System.Drawing.Point(6, 6)
        self._lbRegionsMain.Name = "lbRegionsMain"
        self._lbRegionsMain.SelectionMode = System.Windows.Forms.SelectionMode.MultiExtended
        self._lbRegionsMain.Size = System.Drawing.Size(192, 381)
        self._lbRegionsMain.TabIndex = 1
        self._lbRegionsMain.SelectedIndexChanged += self.LbRegions2SelectedIndexChanged
        #
        # lbPeopleTypes
        #
        self._lbPeopleTypes.Anchor = System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left
        self._lbPeopleTypes.FormattingEnabled = True
        self._lbPeopleTypes.Location = System.Drawing.Point(6, 397)
        self._lbPeopleTypes.Name = "lbPeopleTypes"
        self._lbPeopleTypes.SelectionMode = System.Windows.Forms.SelectionMode.MultiExtended
        self._lbPeopleTypes.Size = System.Drawing.Size(192, 186)
        self._lbPeopleTypes.TabIndex = 2
        self._lbPeopleTypes.SelectedIndexChanged += self.LbPeopleTypesSelectedIndexChanged
        #
        # ID
        #
        self._ID.HeaderText = "ID"
        self._ID.Name = "ID"
        self._ID.ReadOnly = True
        self._ID.Visible = False
        #
        # Group
        #
        self._Group.HeaderText = "Group"
        self._Group.Name = "Group"
        self._Group.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.Automatic
        #
        # Area
        #
        self._Area.HeaderText = "Area"
        self._Area.Name = "Area"
        self._Area.ReadOnly = True
        #
        # clFName
        #
        self._clFName.HeaderText = "First Name"
        self._clFName.Name = "clFName"
        self._clFName.Width = 150
        #
        # clLName
        #
        self._clLName.HeaderText = "Last Name"
        self._clLName.Name = "clLName"
        self._clLName.Width = 150
        #
        # clPhone
        #
        self._clPhone.HeaderText = "Phone Number"
        self._clPhone.Name = "clPhone"
        #
        # clAddress
        #
        self._clAddress.HeaderText = "Address"
        self._clAddress.Name = "clAddress"
        self._clAddress.Width = 350
        #
        # Comments
        #
        self._Comments.HeaderText = "Comments"
        self._Comments.Name = "Comments"
        #
        # Latitude
        #
        self._Latitude.HeaderText = "Latitude"
        self._Latitude.Name = "Latitude"
        #
        # Longitude
        #
        self._Longitude.HeaderText = "Longitude"
        self._Longitude.Name = "Longitude"
        #
        # btnDelete
        #
        self._btnDelete.Anchor = System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left
        self._btnDelete.Location = System.Drawing.Point(3, 560)
        self._btnDelete.Name = "btnDelete"
        self._btnDelete.Size = System.Drawing.Size(59, 23)
        self._btnDelete.TabIndex = 4
        self._btnDelete.Text = "Delete"
        self._btnDelete.UseVisualStyleBackColor = True
        self._btnDelete.Click += self.BtnRegionDeleteClick
        #
        # MainForm
        #
        self.ClientSize = System.Drawing.Size(984, 612)
        self.Controls.Add(self._tcMain)
        self.Name = "MainForm"
        self.Text = "EAB - Electronic Area Book"
        self.FormClosing += self.FormClosingCleanup
        self._tcMain.ResumeLayout(False)
        self._tpPeople.ResumeLayout(False)
        self._tpPeople.PerformLayout()
        self._dgvPeople.EndInit()
        self._tpMap.ResumeLayout(False)
        self._tpRegions.ResumeLayout(False)
        self.ResumeLayout(False)

    def InitializeData( self ):
        self._rowHighlight = System.Windows.Forms.DataGridViewCellStyle()

        self._rowHighlight.BackColor          = System.Drawing.Color.FromArgb( 67, 107, 149 )
        self._rowHighlight.SelectionBackColor = System.Drawing.Color.FromArgb( 67, 107, 149 )
        self._rowHighlight.ForeColor          = System.Drawing.Color.FromArgb( 255, 255, 255 )
        self._rowHighlight.SelectionForeColor = System.Drawing.Color.FromArgb( 255, 255, 255 )

        self._rowDefault = System.Windows.Forms.DataGridViewCellStyle()

        self._rowDefault.BackColor          = System.Drawing.Color.FromArgb( 255, 255, 255 )
        self._rowDefault.SelectionBackColor = System.Drawing.Color.FromArgb( 255, 255, 255 )
        self._rowDefault.ForeColor          = System.Drawing.Color.FromArgb( 0, 0, 0 )
        self._rowDefault.SelectionForeColor = System.Drawing.Color.FromArgb( 0, 0, 0 )

        people = self.dbHelper.getPeople()

        if people != None:
            self.DGVAddPeople( people )

        types = self.dbHelper.getGroups()
        if types != None and len(types) > 0:
            for t in types:
                self._lbPeopleTypes.Items.Add( t[1] )
            self._lbPeopleTypes.SelectedItem = types[0][1]

        areas = self.dbHelper.getAreas()

        if areas != None and len(areas) > 0:
            self.refreshAreas()
        else:
            self._lbRegionsMain.Items.Add( 'None' )
            self._lbRegionsMain.SelectedItem = 'None'

        # self.mapHelperMain.buildPage( self.dbHelper.getAreas(), self.dbHelper.getPeople() )
        self.RedrawMainMap()

    def refreshAreas( self ):
        self._lbRegions    .Items.Clear()
        self._lbRegionsMain.Items.Clear()

        self._lbRegionsMain.Items.Add( 'None' )
        self._lbRegionsMain.SelectedItem = 'None'

        areas = self.dbHelper.getAreas()
        if areas != None and len(areas) > 0:
            areas = sorted( areas, key = itemgetter( 1 ) )

            for area in areas:
                self._lbRegions    .Items.Add( area[1] )
                self._lbRegionsMain.Items.Add( area[1] )

            self.mapHelperRegionCurrent = areas[0][1]
            self.mapHelperRegion.buildPage( areas[0], self.mapHelperRegion.getRegionPointsAsPeople( areas[0] ) )

            self._lbRegions.SelectedItem = areas[0][1]

        else:
            self.mapHelperRegion.buildPage()

    def LbRegionsSelectedIndexChanged(self, sender, e):
        # Try to prevent extra page building
        if self.mapHelperRegionCurrent != self._lbRegions.SelectedItem:
            area = self.dbHelper.getArea( self._lbRegions.SelectedItem )
            points = self.mapHelperRegion.getRegionPointsAsPeople( area )
            self.mapHelperRegion.buildPage( area, points )
            self.mapHelperRegionCurrent = self._lbRegions.SelectedItem

    def LbPeopleTypesSelectedIndexChanged(self, sender, e):
        self.RedrawMainMap()

    def LbRegions2SelectedIndexChanged(self, sender, e):
        self.RedrawMainMap()

    def RedrawMainMap( self ):
        areas = [ self.dbHelper.getArea( item ) for item in self._lbRegionsMain.SelectedItems ]

        people = list()

        if areas:
            for person in self.dbHelper.getPeopleFromTypesInAreas( [item for item in self._lbPeopleTypes.SelectedItems], areas ):
                people.append( person )

        self.mapHelperMain.buildPage( areas, people )

    def BtnAddPeopleClick(self, sender, e):
        form = AddPeopleForm( self )
        form.ShowDialog()

    def BtnRegionNewClick(self, sender, e):
        form = AddRegionForm( self )
        form.ShowDialog()

    def BtnRegionEditClick(self, sender, e):
        if self._lbRegions.SelectedItem != None:
            form = AddRegionForm( self, self._lbRegions.SelectedItem )
            form.ShowDialog()

    def BtnRegionDeleteClick(self, sender, e):
        if self._lbRegions.SelectedItem != None:
            result = MessageBox.Show( "Are you sure that you want to delete '%s'?\nIt will be gone forever..." % self._lbRegions.SelectedItem, "Warning", MessageBoxButtons.YesNo, MessageBoxIcon.Question )
            if result == DialogResult.Yes:
                area = self.dbHelper.getArea( self._lbRegions.SelectedItem )
                self.dbHelper.removeArea( area[0] )
                self.refreshAreas()

    def TbSearchKeyUp(self, sender, e):
        count = 0

        for row in self._dgvPeople.Rows:
            inRow = False

            for cell in row.Cells:
                try:
                    if unicode(sender.Text).lower() in unicode(cell.Value).lower():
                        count += 1
                        inRow = True
                        break
                except Exception, e:
                    print 'Search Error:', e
                    raise e

            row.Visible = inRow

        self._tbPeopleCount.Text = str(count)

    def FormClosingCleanup(self, sender, e):
        self.mapHelperMain.cleanup()
        self.mapHelperRegion.cleanup()
        self.dbHelper.close()

    # ===========================================================
    # DataGridView
    # ===========================================================
    def DgvPeopleRowEnter(self, sender, e):
        for cell in self._dgvPeople.Rows[e.RowIndex].Cells:
            if cell.ColumnIndex in ( 0, 1 ): continue
            cell.Style = self._rowHighlight

    def DgvPeopleRowLeave(self, sender, e):
        for cell in self._dgvPeople.Rows[e.RowIndex].Cells:
            if cell.ColumnIndex in ( 0, 1 ): continue
            cell.Style = self._rowDefault

    def DgvPeopleCellClick(self, sender, e):
        if e.ColumnIndex == 1: # The Combo Box Cell
            self._dgvPeople.Rows[e.RowIndex].Selected = True
            System.Windows.Forms.SendKeys.Send( "{f4}" )

    def DgvPeopleCellEndEdit(self, sender, e):
        success = self.dbHelper.updatePerson(
            self._dgvPeople.Rows[e.RowIndex].Cells[0].Value,
            self._dgvPeople.Rows[e.RowIndex].Cells[1].Value,
            # self._dgvPeople.Rows[e.RowIndex].Cells[2].Value, # Changed if lat,lng changed
            self._dgvPeople.Rows[e.RowIndex].Cells[3].Value,
            self._dgvPeople.Rows[e.RowIndex].Cells[4].Value,
            self._dgvPeople.Rows[e.RowIndex].Cells[5].Value,
            self._dgvPeople.Rows[e.RowIndex].Cells[6].Value,
            self._dgvPeople.Rows[e.RowIndex].Cells[7].Value,
            float(self._dgvPeople.Rows[e.RowIndex].Cells[8].Value),
            float(self._dgvPeople.Rows[e.RowIndex].Cells[9].Value),
        )

        # if not success:
            # MessageBox.Show( 'Your address was invalid!\nTry again...', 'Bad!', MessageBoxButtons.OK, MessageBoxIcon.Error )
            # self._dgvPeople.Rows[e.RowIndex].ErrorText = 'Your address was invalid!\nTry again...'
            # return

        person = self.dbHelper.getPerson( **{'p.id': self._dgvPeople.Rows[e.RowIndex].Cells[0].Value} )

        self._dgvPeople.Rows[e.RowIndex].Cells[1].Value = person[2]
        self._dgvPeople.Rows[e.RowIndex].Cells[2].Value = person[1]
        self._dgvPeople.Rows[e.RowIndex].Cells[3].Value = person[3]
        self._dgvPeople.Rows[e.RowIndex].Cells[4].Value = person[4]
        self._dgvPeople.Rows[e.RowIndex].Cells[5].Value = person[5]
        self._dgvPeople.Rows[e.RowIndex].Cells[6].Value = person[6]
        self._dgvPeople.Rows[e.RowIndex].Cells[7].Value = person[7]
        self._dgvPeople.Rows[e.RowIndex].Cells[8].Value = person[8]
        self._dgvPeople.Rows[e.RowIndex].Cells[9].Value = person[9]

    def DgvPeopleUserDeletingRow(self, sender, e):
        result = MessageBox.Show( "Are you sure that you want to delete %s %s?\nThey will be gone forever..." % ( e.Row.Cells[3].Value, e.Row.Cells[4].Value ), "Warning", MessageBoxButtons.YesNo, MessageBoxIcon.Question )
        if result == DialogResult.No:
            e.Cancel = True

    def DgvPeopleUserDeletedRow(self, sender, e):
        self.dbHelper.removePerson( e.Row.Cells[0].Value )

    def DgvPeopleCellValidating(self, sender, e):

        if e.FormattedValue == sender.Rows[e.RowIndex].Cells[e.ColumnIndex].Value:
            return

        if e.ColumnIndex == 8 or e.ColumnIndex == 9:
            try:
                float( e.FormattedValue )
                sender.Rows[e.RowIndex].ErrorText = ""
            except:
                e.Cancel = True
                sender.Rows[e.RowIndex].ErrorText = "Must be a propper Longitude or Latittude value! (... A number)"

        if e.ColumnIndex == 6:
            info = self.dbHelper.geocoder.geocode( e.FormattedValue )

            if info == None:
                # e.Cancel = True
                # sender.Rows[e.RowIndex].Cells[6].Value = e.FormattedValue
                sender.Rows[e.RowIndex].ErrorText = "Your address is crap! Try again."
            else:
                sender.Rows[e.RowIndex].Cells[8].Value = info.latitude
                sender.Rows[e.RowIndex].Cells[9].Value = info.longitude
                sender.Rows[e.RowIndex].ErrorText = ""


    def DGVAddPeople( self, people, clear=True ):
        groups = self.dbHelper.getGroups()

        if clear: self._dgvPeople.Rows.Clear()

        for person in people:

            self.DGVAddPerson( person, groups )

    def DGVAddPerson( self, person, groups=None ):
        groups = self.dbHelper.getGroups() if groups == None else groups

        index = self._dgvPeople.Rows.Count

        self._dgvPeople.Rows.Add()

        self._dgvPeople.Rows[ index ].Cells[0].Value = person[0]
        self._dgvPeople.Rows[ index ].Cells[2].Value = person[1]
        self._dgvPeople.Rows[ index ].Cells[3].Value = person[3]
        self._dgvPeople.Rows[ index ].Cells[4].Value = person[4]
        self._dgvPeople.Rows[ index ].Cells[5].Value = person[5]
        self._dgvPeople.Rows[ index ].Cells[6].Value = person[6]
        self._dgvPeople.Rows[ index ].Cells[7].Value = person[7]
        self._dgvPeople.Rows[ index ].Cells[8].Value = person[8]
        self._dgvPeople.Rows[ index ].Cells[9].Value = person[9]

        if 'Bad Address' in person[6] or person[8] == 0.0 or person[9] == 0.0:
            self._dgvPeople.Rows[ index ].ErrorText = 'There is a bad address here!'

        for g in groups:
            self._dgvPeople.Rows[ index ].Cells[1].Items.Add( g[1] )

        self._dgvPeople.Rows[ index ].Cells[1].Value = str(person[2])