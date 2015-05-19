# 1.1
import System.Drawing
import System.Windows.Forms

from System.Drawing import *
from System.Windows.Forms import *
from System import Environment

from DatabaseHelper import DatabaseHelper

class AddPeopleForm(Form):
    def __init__(self, parent):
        self._parent = parent
        self.InitializeComponent()

        self.initializeBatchHelp()

        self.loadGroups()

        if len( self._cbGroup.Items) > 0:
            self._cbGroup.SelectedIndex = 0

    def InitializeComponent(self):
        self._tcPeople = System.Windows.Forms.TabControl()
        self._tpAddIndivudual = System.Windows.Forms.TabPage()
        self._tpAddBatch = System.Windows.Forms.TabPage()
        self._btnBatchSubmit = System.Windows.Forms.Button()
        self._btnIndividualSubmit = System.Windows.Forms.Button()
        self._tbComments = System.Windows.Forms.TextBox()
        self._tbPhone = System.Windows.Forms.TextBox()
        self._tbAddress = System.Windows.Forms.TextBox()
        self._tbLastName = System.Windows.Forms.TextBox()
        self._tbFirstName = System.Windows.Forms.TextBox()
        self._lblAddPersonComments = System.Windows.Forms.Label()
        self._lblAddPersonPhoneNumber = System.Windows.Forms.Label()
        self._lblAddPersonAddress = System.Windows.Forms.Label()
        self._lblAddPersonLastName = System.Windows.Forms.Label()
        self._lblAddPersonFirstName = System.Windows.Forms.Label()
        self._cbGroup = System.Windows.Forms.ComboBox()
        self._lblGroup = System.Windows.Forms.Label()
        self._btnAddGroup = System.Windows.Forms.Button()
        self._tbGroupName = System.Windows.Forms.TextBox()
        self._lblAddGroup = System.Windows.Forms.Label()
        self._lblGroupAbreviation = System.Windows.Forms.Label()
        self._tbGroupAbreviation = System.Windows.Forms.TextBox()
        self._lblBatchInfo = System.Windows.Forms.Label()
        self._pbBatch = System.Windows.Forms.ProgressBar()
        self._bgBatch = System.ComponentModel.BackgroundWorker()
        self._rtbBatch = System.Windows.Forms.RichTextBox()
        self._gbAddGroup = System.Windows.Forms.GroupBox()
        self._tcPeople.SuspendLayout()
        self._tpAddIndivudual.SuspendLayout()
        self._tpAddBatch.SuspendLayout()
        self._gbAddGroup.SuspendLayout()
        self.SuspendLayout()
        #
        # tcPeople
        #
        self._tcPeople.Controls.Add(self._tpAddIndivudual)
        self._tcPeople.Controls.Add(self._tpAddBatch)
        self._tcPeople.Dock = System.Windows.Forms.DockStyle.Fill
        self._tcPeople.Location = System.Drawing.Point(0, 0)
        self._tcPeople.Name = "tcPeople"
        self._tcPeople.SelectedIndex = 0
        self._tcPeople.Size = System.Drawing.Size(984, 612)
        self._tcPeople.TabIndex = 0
        #
        # tpAddIndivudual
        #
        self._tpAddIndivudual.Controls.Add(self._gbAddGroup)
        self._tpAddIndivudual.Controls.Add(self._lblGroup)
        self._tpAddIndivudual.Controls.Add(self._cbGroup)
        self._tpAddIndivudual.Controls.Add(self._btnIndividualSubmit)
        self._tpAddIndivudual.Controls.Add(self._tbComments)
        self._tpAddIndivudual.Controls.Add(self._tbPhone)
        self._tpAddIndivudual.Controls.Add(self._tbAddress)
        self._tpAddIndivudual.Controls.Add(self._tbLastName)
        self._tpAddIndivudual.Controls.Add(self._tbFirstName)
        self._tpAddIndivudual.Controls.Add(self._lblAddPersonComments)
        self._tpAddIndivudual.Controls.Add(self._lblAddPersonPhoneNumber)
        self._tpAddIndivudual.Controls.Add(self._lblAddPersonAddress)
        self._tpAddIndivudual.Controls.Add(self._lblAddPersonLastName)
        self._tpAddIndivudual.Controls.Add(self._lblAddPersonFirstName)
        self._tpAddIndivudual.Location = System.Drawing.Point(4, 22)
        self._tpAddIndivudual.Name = "tpAddIndivudual"
        self._tpAddIndivudual.Padding = System.Windows.Forms.Padding(3)
        self._tpAddIndivudual.Size = System.Drawing.Size(976, 586)
        self._tpAddIndivudual.TabIndex = 0
        self._tpAddIndivudual.Text = "Single Person"
        self._tpAddIndivudual.UseVisualStyleBackColor = True
        #
        # tpAddBatch
        #
        self._tpAddBatch.Controls.Add(self._rtbBatch)
        self._tpAddBatch.Controls.Add(self._pbBatch)
        self._tpAddBatch.Controls.Add(self._lblBatchInfo)
        self._tpAddBatch.Controls.Add(self._btnBatchSubmit)
        self._tpAddBatch.Location = System.Drawing.Point(4, 22)
        self._tpAddBatch.Name = "tpAddBatch"
        self._tpAddBatch.Padding = System.Windows.Forms.Padding(3)
        self._tpAddBatch.Size = System.Drawing.Size(1038, 537)
        self._tpAddBatch.TabIndex = 1
        self._tpAddBatch.Text = "Batch (Advanced)"
        self._tpAddBatch.UseVisualStyleBackColor = True
        #
        # btnBatchSubmit
        #
        self._btnBatchSubmit.Anchor = System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right
        self._btnBatchSubmit.Location = System.Drawing.Point(919, 504)
        self._btnBatchSubmit.Name = "btnBatchSubmit"
        self._btnBatchSubmit.Size = System.Drawing.Size(111, 30)
        self._btnBatchSubmit.TabIndex = 1
        self._btnBatchSubmit.Text = "Submit"
        self._btnBatchSubmit.UseVisualStyleBackColor = True
        self._btnBatchSubmit.Click += self.BtnBatchSubmitClick
        #
        # btnIndividualSubmit
        #
        self._btnIndividualSubmit.Anchor = System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left
        self._btnIndividualSubmit.Location = System.Drawing.Point(9, 552)
        self._btnIndividualSubmit.Name = "btnIndividualSubmit"
        self._btnIndividualSubmit.Size = System.Drawing.Size(105, 26)
        self._btnIndividualSubmit.TabIndex = 21
        self._btnIndividualSubmit.Text = "Save"
        self._btnIndividualSubmit.UseVisualStyleBackColor = True
        self._btnIndividualSubmit.Click += self.BtnIndividualSubmitClick
        #
        # tbComments
        #
        self._tbComments.Anchor = System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left | System.Windows.Forms.AnchorStyles.Right
        self._tbComments.Location = System.Drawing.Point(126, 111)
        self._tbComments.Multiline = True
        self._tbComments.Name = "tbComments"
        self._tbComments.Size = System.Drawing.Size(841, 467)
        self._tbComments.TabIndex = 20
        #
        # tbPhone
        #
        self._tbPhone.Anchor = System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left | System.Windows.Forms.AnchorStyles.Right
        self._tbPhone.Location = System.Drawing.Point(126, 85)
        self._tbPhone.Name = "tbPhone"
        self._tbPhone.Size = System.Drawing.Size(841, 20)
        self._tbPhone.TabIndex = 19
        #
        # tbAddress
        #
        self._tbAddress.Anchor = System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left | System.Windows.Forms.AnchorStyles.Right
        self._tbAddress.Location = System.Drawing.Point(126, 59)
        self._tbAddress.Name = "tbAddress"
        self._tbAddress.Size = System.Drawing.Size(841, 20)
        self._tbAddress.TabIndex = 18
        #
        # tbLastName
        #
        self._tbLastName.Anchor = System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left | System.Windows.Forms.AnchorStyles.Right
        self._tbLastName.Location = System.Drawing.Point(126, 33)
        self._tbLastName.Name = "tbLastName"
        self._tbLastName.Size = System.Drawing.Size(841, 20)
        self._tbLastName.TabIndex = 17
        #
        # tbFirstName
        #
        self._tbFirstName.Anchor = System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left | System.Windows.Forms.AnchorStyles.Right
        self._tbFirstName.Location = System.Drawing.Point(126, 7)
        self._tbFirstName.Name = "tbFirstName"
        self._tbFirstName.Size = System.Drawing.Size(841, 20)
        self._tbFirstName.TabIndex = 16
        #
        # lblAddPersonComments
        #
        self._lblAddPersonComments.Location = System.Drawing.Point(26, 111)
        self._lblAddPersonComments.Name = "lblAddPersonComments"
        self._lblAddPersonComments.Size = System.Drawing.Size(94, 20)
        self._lblAddPersonComments.TabIndex = 15
        self._lblAddPersonComments.Text = "Comments:"
        self._lblAddPersonComments.TextAlign = System.Drawing.ContentAlignment.TopRight
        #
        # lblAddPersonPhoneNumber
        #
        self._lblAddPersonPhoneNumber.Location = System.Drawing.Point(20, 85)
        self._lblAddPersonPhoneNumber.Name = "lblAddPersonPhoneNumber"
        self._lblAddPersonPhoneNumber.Size = System.Drawing.Size(100, 20)
        self._lblAddPersonPhoneNumber.TabIndex = 14
        self._lblAddPersonPhoneNumber.Text = "Phone Number:"
        self._lblAddPersonPhoneNumber.TextAlign = System.Drawing.ContentAlignment.TopRight
        #
        # lblAddPersonAddress
        #
        self._lblAddPersonAddress.Location = System.Drawing.Point(20, 59)
        self._lblAddPersonAddress.Name = "lblAddPersonAddress"
        self._lblAddPersonAddress.Size = System.Drawing.Size(100, 20)
        self._lblAddPersonAddress.TabIndex = 13
        self._lblAddPersonAddress.Text = "Address:"
        self._lblAddPersonAddress.TextAlign = System.Drawing.ContentAlignment.TopRight
        #
        # lblAddPersonLastName
        #
        self._lblAddPersonLastName.Location = System.Drawing.Point(20, 33)
        self._lblAddPersonLastName.Name = "lblAddPersonLastName"
        self._lblAddPersonLastName.Size = System.Drawing.Size(100, 20)
        self._lblAddPersonLastName.TabIndex = 12
        self._lblAddPersonLastName.Text = "Last Name:"
        self._lblAddPersonLastName.TextAlign = System.Drawing.ContentAlignment.TopRight
        #
        # lblAddPersonFirstName
        #
        self._lblAddPersonFirstName.Location = System.Drawing.Point(17, 7)
        self._lblAddPersonFirstName.Name = "lblAddPersonFirstName"
        self._lblAddPersonFirstName.Size = System.Drawing.Size(103, 20)
        self._lblAddPersonFirstName.TabIndex = 11
        self._lblAddPersonFirstName.Text = "First Name:"
        self._lblAddPersonFirstName.TextAlign = System.Drawing.ContentAlignment.TopRight
        #
        # cbGroup
        #
        self._cbGroup.FormattingEnabled = True
        self._cbGroup.Location = System.Drawing.Point(6, 147)
        self._cbGroup.Name = "cbGroup"
        self._cbGroup.Size = System.Drawing.Size(114, 21)
        self._cbGroup.TabIndex = 22
        #
        # lblGroup
        #
        self._lblGroup.Location = System.Drawing.Point(6, 131)
        self._lblGroup.Name = "lblGroup"
        self._lblGroup.Size = System.Drawing.Size(100, 13)
        self._lblGroup.TabIndex = 23
        self._lblGroup.Text = "Group:"
        #
        # btnAddGroup
        #
        self._btnAddGroup.Location = System.Drawing.Point(6, 99)
        self._btnAddGroup.Name = "btnAddGroup"
        self._btnAddGroup.Size = System.Drawing.Size(105, 24)
        self._btnAddGroup.TabIndex = 24
        self._btnAddGroup.Text = "Add"
        self._btnAddGroup.UseVisualStyleBackColor = True
        self._btnAddGroup.Click += self.BtnAddGroupClick
        #
        # tbGroupName
        #
        self._tbGroupName.Location = System.Drawing.Point(6, 33)
        self._tbGroupName.Name = "tbGroupName"
        self._tbGroupName.Size = System.Drawing.Size(105, 20)
        self._tbGroupName.TabIndex = 25
        #
        # lblAddGroup
        #
        self._lblAddGroup.Location = System.Drawing.Point(3, 16)
        self._lblAddGroup.Name = "lblAddGroup"
        self._lblAddGroup.Size = System.Drawing.Size(100, 14)
        self._lblAddGroup.TabIndex = 26
        self._lblAddGroup.Text = "Group Name:"
        #
        # lblGroupAbreviation
        #
        self._lblGroupAbreviation.Location = System.Drawing.Point(3, 56)
        self._lblGroupAbreviation.Name = "lblGroupAbreviation"
        self._lblGroupAbreviation.Size = System.Drawing.Size(100, 14)
        self._lblGroupAbreviation.TabIndex = 28
        self._lblGroupAbreviation.Text = "Group Abreviation:"
        #
        # tbGroupAbreviation
        #
        self._tbGroupAbreviation.Location = System.Drawing.Point(6, 73)
        self._tbGroupAbreviation.MaxLength = 2
        self._tbGroupAbreviation.Name = "tbGroupAbreviation"
        self._tbGroupAbreviation.Size = System.Drawing.Size(105, 20)
        self._tbGroupAbreviation.TabIndex = 27
        #
        # lblBatchInfo
        #
        self._lblBatchInfo.Anchor = System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left
        self._lblBatchInfo.Font = System.Drawing.Font("Microsoft Sans Serif", 8.25, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 0)
        self._lblBatchInfo.Location = System.Drawing.Point(6, 3)
        self._lblBatchInfo.Name = "lblBatchInfo"
        self._lblBatchInfo.Size = System.Drawing.Size(161, 528)
        self._lblBatchInfo.TabIndex = 2
        self._lblBatchInfo.Text = "Batch info:"
        #
        # pbBatch
        #
        self._pbBatch.Anchor = System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left | System.Windows.Forms.AnchorStyles.Right
        self._pbBatch.Location = System.Drawing.Point(173, 508)
        self._pbBatch.Name = "pbBatch"
        self._pbBatch.Size = System.Drawing.Size(740, 23)
        self._pbBatch.TabIndex = 3
        self._pbBatch.Visible = False
        #
        # bgBatch
        #
        self._bgBatch.WorkerReportsProgress = True
        self._bgBatch.DoWork += self.BgBatchDoWork
        self._bgBatch.ProgressChanged += self.BgBatchProgressChanged
        self._bgBatch.RunWorkerCompleted += self.BgBatchRunWorkerCompleted
        #
        # rtbBatch
        #
        self._rtbBatch.Anchor = System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left | System.Windows.Forms.AnchorStyles.Right
        self._rtbBatch.Font = System.Drawing.Font("Courier New", 8.25, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 0)
        self._rtbBatch.Location = System.Drawing.Point(173, 6)
        self._rtbBatch.Name = "rtbBatch"
        self._rtbBatch.Size = System.Drawing.Size(857, 496)
        self._rtbBatch.TabIndex = 4
        self._rtbBatch.Text = ""
        #
        # gbAddGroup
        #
        self._gbAddGroup.Controls.Add(self._btnAddGroup)
        self._gbAddGroup.Controls.Add(self._lblGroupAbreviation)
        self._gbAddGroup.Controls.Add(self._tbGroupName)
        self._gbAddGroup.Controls.Add(self._tbGroupAbreviation)
        self._gbAddGroup.Controls.Add(self._lblAddGroup)
        self._gbAddGroup.Location = System.Drawing.Point(3, 174)
        self._gbAddGroup.Name = "gbAddGroup"
        self._gbAddGroup.Size = System.Drawing.Size(117, 133)
        self._gbAddGroup.TabIndex = 29
        self._gbAddGroup.TabStop = False
        self._gbAddGroup.Text = "Add New Group"
        #
        # AddPeopleForm
        #
        self.ClientSize = System.Drawing.Size(984, 612)
        self.Controls.Add(self._tcPeople)
        self.Name = "AddPeopleForm"
        self.Text = "Add People To Database"
        self._tcPeople.ResumeLayout(False)
        self._tpAddIndivudual.ResumeLayout(False)
        self._tpAddIndivudual.PerformLayout()
        self._tpAddBatch.ResumeLayout(False)
        self._gbAddGroup.ResumeLayout(False)
        self._gbAddGroup.PerformLayout()
        self.ResumeLayout(False)

    def initializeBatchHelp( self ):
        n = Environment.NewLine

        self._lblBatchInfo.Text = \
            "Batch info:" + n + \
            "Input CSV styled text into the right box. The first line should contain column headers defining how the CSV is setup." + n * 2 + \
            "Recognized Headers:" + n + \
            "  Group*" + n + \
            "  FirstName" + n + \
            "  LastName" + n + \
            "  CombinedName**" + n + \
            "  Phone" + n + \
            "  Address" + n + \
            "  Comments" + n + \
            "  AltPhone1***" + n + \
            "  AltPhone2***" + n + \
            "  Child#****" + n * 3 + \
            "* You need a Group Column. Also, the groups need to exist in the 'Add Indivual' page." + n * 2 + \
            "** If used, FirstName and LastName will be ignored." + n * 2 + \
            "*** Alternate fields to check for a phone number in the event 'Phone' doesn't have a value." + n * 2 + \
            "**** For Children names, eg: 'Child1', 'Child2', ect..." + n * 2 + \
            "Make sure that all the Groups are added the the list on the Add Individuals page" + n * 2 + \
            "Make suere there are no spaces between the Comma's"

        self._rtbBatch.Text = "Group,FirstName,LastName,Phone,Address,Comments,Child1,Child2" + n + \
                              '"LA","Luke","Skywalker","123 456 7980","Mount Everest, China","This guy kissed his sister.",,' + n + \
                              '"Less Active","Anikan","Skywalker","098 765 4321","The Bermuda Triangle,","A bad dude.","Luke Skywalker","Leah",'


    # ==================================================================
    # Individual Page
    # ==================================================================
    def loadGroups( self ):
        groups = self._parent.dbHelper.getGroups()

        for group in groups:
            self._cbGroup.Items.Add( group[1] )

    def BtnIndividualSubmitClick(self, sender, e):
        # Add to db
        self._parent.dbHelper.addIndividual(
            str(self._cbGroup.SelectedItem.ToString()),
            str(self._tbFirstName.Text),
            str(self._tbLastName.Text),
            str(self._tbPhone.Text),
            str(self._tbAddress.Text),
            str(self._tbComments.Text)
        )

        # Add to DGV
        person = self._parent.dbHelper.getPerson( fname=self._tbFirstName.Text, lname=self._tbLastName.Text )

        self._parent.DGVAddPerson(person)

        self.Close()

    def BtnAddGroupClick(self, sender, e):
        for item in self._cbGroup.Items:
            if self._tbGroupName.Text in item:
                return None

        if self._tbGroupAbreviation.Text == '':
            MessageBox.Show( self, "Need to have an abreviation for your group.", "Whoops" )
            return

        self._cbGroup.Items.Add( self._tbGroupName.Text )
        self._parent.dbHelper.addGroup( self._tbGroupName.Text, self._tbGroupAbreviation.Text )

        self._cbGroup.SelectedIndex = self._cbGroup.FindString( self._tbGroupName.Text )

        self._tbGroupName.Text        = ""
        self._tbGroupAbreviation.Text = ""

    # ==================================================================
    # Batch Background Worker
    # ==================================================================
    def BtnBatchSubmitClick(self, sender, e):
        self._pbBatch.Visible = True
        self._parent.dbHelper.close()
        self._bgBatch.RunWorkerAsync()
        # MessageBox( "Batch process is complete!", "Yay!" )

    def BgBatchDoWork(self, sender, e):
        try:
            # We need a seperate instance of the DBHelper, because it's a different thread
            dbHelper = DatabaseHelper()
            dbHelper.addBatch( self._rtbBatch.Text, self )
        except Exception, e:
            print 'Batch Error: ', e

    def BgBatchProgressChanged(self, sender, e):
        # Update the progress bar
        self._pbBatch.Value = int(e.ProgressPercentage)

    def BgBatchRunWorkerCompleted(self, sender, e):
        self._pbBatch.Visible = False
        self._parent.dbHelper = DatabaseHelper()
        self._rtbBatch.Text = ""

        people = self._parent.dbHelper.getPeople()

        if people != None:
            self._parent.DGVAddPeople( people )

        self.Close()
