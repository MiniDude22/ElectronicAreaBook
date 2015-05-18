# 1.0
# coding=utf-8

# A verry crude autoupdater based off of the github raw files.

import os
import urllib2

directory = os.path.dirname(os.path.abspath(__file__))

mainFiles = (
    'AddPeopleForm.py',
    'AddRegionForm.py',
    'DatabaseHelper.py',
    'MainForm.py',
    'MapHelper.py'
)

sensitiveFiles = (
    'AutoUpdater.py',
    'Program.py',
)

def updateFiles():
    for f in mainFiles:
        updateFile( f )

    sensitive = False
    for f in sensitiveFiles:
        if updateFile( f ):
            sensitive = True

    return sensitive

def updateFile( name ):
    with open( directory + '/' + name, 'r+b' ) as f:
        curVersion = getVersion( f.read(15) ) # only need the first few characters
        newVersion = getVersion( getGitHubPage( name ) )

        if curVersion < newVersion:
            print "Updating '{}' to version {}...".format( name, newVersion )
            f.seek(0)
            f.write( getGitHubPage( name, False ) )
            f.truncate()
            f.flush()
            return True
        else:
            print "'{}' is up to date!".format( name )
            return False

def getVersion( text ):
    try:
        return float( text.split("\n")[0].strip( u'ï»¿# ' ) )
    except Exception, e:
        print 'Invalid version number: "{}"'.format( text )
        return 0.0

def getGitHubPage( fileName, partial=True ):
    url = "https://raw.githubusercontent.com/MiniDude22/ElectronicAreaBook/master/{}".format( fileName )

    if partial:
        req = urllib2.Request( url )
        req.headers['Range'] = 'bytes=%s-%s' % (0, 15) # only get the first few characters of the file
        response = urllib2.urlopen( req )
    else:
        response = urllib2.urlopen( url )

    return response.read()
