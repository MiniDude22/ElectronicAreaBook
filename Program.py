# 1.0

import AutoUpdater

sensitive = AutoUpdater.updateFiles()

if sensitive == False:
    import clr
    clr.AddReference('System.Windows.Forms')
    clr.AddReference('System.Drawing')

    from System.Windows.Forms import Application
    import MainForm

    Application.EnableVisualStyles()
    form = MainForm.MainForm()
    Application.Run(form)

else:
    print 'Sensitive files were updated. Restart the program.'
