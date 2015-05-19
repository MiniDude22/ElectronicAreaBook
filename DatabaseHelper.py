# 1.1
import csv
import json
import os
import sqlite3

from System.Windows.Forms import MessageBox

from GeocoderHelper import GeocoderHelper

class DatabaseHelper:

    def __init__(self):
        self.connection = sqlite3.connect( '{}/People.db'.format( os.path.dirname(os.path.abspath(__file__)) ))

        self.db = self.connection.cursor()
        self.geocoder = GeocoderHelper()

        # self.dropTables()
        self.initializeTables()

    # ==================================================================
    # Tables
    # ==================================================================
    def dropTables( self ):
        try:
            self.db.execute("DROP TABLE Area;"      )
            self.db.execute("DROP TABLE PeopleType;")
            self.db.execute("DROP TABLE People;"    )
        except Exception as e:
            pass

    # If the tables doen't exist, create them.
    def initializeTables( self ):

        tables = {
            'People'     : "CREATE TABLE People     ( id integer primary key, areaID int, typeID int, fName varchar(50), lName varchar(50), phone varchar(20), address varchar(100), comments text, latitude real, longitude real );",
            'Area'       : "CREATE TABLE Area       ( id integer primary key, name varchar(40), sColor varchar(7), sOpacity real, fColor varchar(7), fOpacity real, points text);",
            'PeopleType' : "CREATE TABLE PeopleType ( id integer primary key, name varchar(30), abreviation varchar(30) );",
        }

        for table in ( 'Area', 'PeopleType', 'People' ):

            self.db.execute("SELECT count(*) FROM sqlite_master WHERE type='table' AND name='%s';" % table )

            result = self.db.fetchone()

            if result == ( 0, ):
                self.db.execute( tables[table] )

                if table == 'Area':
                    self.db.execute( 'INSERT INTO Area ( name, sColor, sOpacity, fColor, fOpacity, Points ) VALUES ( ?, ?, ?, ?, ?, ? )', ( 'None', '', 0.0, '', 0.0, '[]' ) )
                elif table == 'PeopleType':
                    self.db.execute( 'INSERT INTO PeopleType ( name, abreviation ) VALUES ( ?, ? )', ( 'None', 'N' ) )

        self.save()

    # ==================================================================
    # Area
    # ==================================================================
    def addArea( self, name, sColor, sOpacity, fColor, fOpacity, points ):
        # If the area doesn't exist create it
        self.db.execute( "SELECT * FROM Area WHERE name = ?;", ( name, ) )
        result = self.db.fetchone()

        if result == None:
            self.db.execute( "INSERT INTO Area ( name, sColor, sOpacity, fColor, fOpacity, Points ) VALUES ( ?, ?, ?, ?, ?, ? );", (
                name,
                sColor,
                float(sOpacity),
                fColor,
                float(fOpacity),
                json.dumps(points)
            ))
            self.save()
            areaID = self.db.lastrowid
        else:
            areaID = result[0]

        return areaID

    def getArea( self, name ):
        self.db.execute( "SELECT * FROM Area WHERE name = ?;", (name,) )

        result = self.db.fetchone()

        if result:
            return result

    def getAreas( self ):
        self.db.execute( "SELECT * FROM Area WHERE name != 'None';" )

        return self.db.fetchall()

    def updateArea( self, ID, name, scolor, sopacity, fcolor, fopacity, points ):
        self.db.execute( """UPDATE Area SET name=?, sColor=?, sOpacity=?, fColor=?, fOpacity=?, points=? WHERE id=?""", (
            name,
            scolor,
            sopacity,
            fcolor,
            fopacity,
            json.dumps(points),
            ID
        ))

        self.save()

    def removeArea( self, ID ):
        self.db.execute( """DELETE FROM Area WHERE id = ?""", ( ID, ) )

        self.save()

    def getAreaRegion( self, name ):
        self.db.execute( "SELECT * FROM Area WHERE name = ?;", name )

        result = self.db.fetchone()

        if result:
            return json.loads( result[5] )

    def getAreaFromPoint( self, latitude, longitude ):
        for area in self.getAreas():

            region = json.loads( area[6] )

            if self.pointInPoly( latitude, longitude, region ):

                return ( area[0], area[1] )

        return ( self.getArea('None')[0], 'None' )

    # ==================================================================
    # Group
    # ==================================================================
    def addGroup( self, name, abreviation ):
        self.db.execute( "INSERT INTO PeopleType ( name, abreviation ) VALUES ( ?, ? );", ( name, abreviation ) )
        self.save()

    def getGroups( self ):
        self.db.execute( "SELECT * FROM PeopleType;" )

        return self.db.fetchall()

    def getGroup( self, name ):
        self.db.execute( "SELECT * FROM PeopleType WHERE name = ? OR abreviation = ?;", (name, name) )

        result = self.db.fetchone()

        if result:
            return result

    # ==================================================================
    # People
    # ==================================================================
    # TODO: Fix Phone numbers for consistency
    def addIndividual( self, group, fname, lname, phone, address, comments ):
        groupID = self.getGroup( group )

        if groupID == None:
            groupID = (0, )

        info = self.geocoder.geocode( address )

        if info == None:
            info = {
                'latitude': 0.0,
                'longitude': 0.0
            }
            address += '(Bad Address)'
        else:
            info = {
                'latitude': info.latitude,
                'longitude': info.longitude
            }

        area = self.getAreaFromPoint( info['latitude'], info['longitude'] )

        self.db.execute( "INSERT INTO People ( areaID, typeID, fName, lName, phone, address, comments, latitude, longitude ) VALUES ( ?, ?, ?, ?, ?, ?, ?, ?, ? );", (
            area[0],
            groupID[0],
            fname,
            lname,
            phone,
            address,
            comments,
            info['latitude'],
            info['longitude']
        ))

        self.save()

    def addBatch( self, batch, parent ):
        try:
            lines = batch.split("\n")

            people = csv.DictReader( lines )

            total = float(len(lines) - 1)
            current = 0

            # results = dict()

            for person in people:
                if 'CombinedName' in person:
                    names = person['CombinedName'].split(',')
                    person['FirstName'] = names[1].strip()
                    person['LastName']  = names[0].strip()

                if person['Phone'] == '':
                    if 'AltPhone1' in person and person['AltPhone1'] != '':
                        person['Phone'] = person['AltPhone1']
                    elif 'AltPhone2' in person and person['AltPhone2'] != '':
                        person['Phone'] = person['AltPhone2']

                children = ', '.join( [ person[header].replace( person['LastName'], '' ).strip( ' ,' ) for header in person if header != None and 'Child' in header and person[header] != '' ] )

                if 'Comments' not in person:
                    person['Comments'] = ''

                if children != '':
                    person['Comments'] += "Children: " + children

                result = self.addIndividual(
                    person['Group'],
                    person['FirstName'],
                    person['LastName'],
                    person['Phone'],
                    person['Address'],
                    person['Comments']
                )

                # results[ person['FirstName'] + ' ' + person['LastName'] ] = result

                current += 1

                parent._bgBatch.ReportProgress( int((current/total) * 100) )
        except Exception, e:
            print 'Add Batch Error:', e
            raise e

        self.close()

    def getPeople( self ):
        self.db.execute( """SELECT p.id, a.name, pt.name, p.fName, p.lName, p.phone, p.address, p.comments, p.latitude, p.longitude, pt.abreviation FROM People AS p
                            LEFT JOIN PeopleType AS pt ON p.typeID = pt.id
                            LEFT JOIN Area AS a ON p.areaID = a.id;""" )

        return self.db.fetchall()

    def getPeopleFromTypesInAreas( self, types, areas ):
        if len(areas) == 0: return None

        self.db.execute( """SELECT p.id, a.name, pt.name, p.fName, p.lName, p.phone, p.address, p.comments, p.latitude, p.longitude, pt.abreviation FROM People AS p
                            LEFT JOIN PeopleType AS pt ON p.typeID = pt.id
                            LEFT JOIN Area AS a ON p.areaID = a.id
                            WHERE pt.name in ( %s )
                            AND p.areaID in ( %s );""" % ( ",".join("?"*len(types)), ",".join("?"*len(areas)) ),
                            list( types ) + [ area[0] for area in areas ] )

        return self.db.fetchall()

    def getPerson( self, **kwargs ):
        self.db.execute( """SELECT p.id, a.name, pt.name, p.fName, p.lName, p.phone, p.address, p.comments, p.latitude, p.longitude, pt.abreviation FROM People AS p
                            LEFT JOIN PeopleType AS pt ON p.typeID = pt.id
                            LEFT JOIN Area AS a ON p.areaID = a.id
                            WHERE %s;""" % ("%s = ? AND "*len(kwargs))[:-5] % tuple( kwargs.keys() ), tuple( kwargs.values() ) )

        result = self.db.fetchone()

        if result: return result

    def removePerson( self, personID ):
        self.db.execute( """DELETE FROM People WHERE id=?""", (personID,) )

        self.save()

    def updatePerson( self, ID, group, fname, lname, phone, address, comments, lat, lng ):
        person = self.getPerson( **{'p.id': ID} )

        area = person[1]

        # print ID, area, group, fname, lname, phone, address, comments, lat, lng

        changed = False

        if address != person[6]:
            changed = True
            info = self.geocoder.geocode( address )

            if info == None: return False

            lat = info.latitude
            lng = info.longitude

        if lat != person[8] or lng != person[9]:
            changed = True
            info = self.getAreaFromPoint( lat, lng )
            area = info[0]

        if group != person[2] or fname != person[3] or lname != person[4] or phone != person[5] or comments != person[7]:
            changed = True

        if changed:
            if isinstance( area,  str ): area  = self.getArea ( area  )[0]
            if isinstance( group, str ): group = self.getGroup( group )[0]

            # print ID, area, group, fname, lname, phone, address, comments, lat, lng

            self.db.execute("""UPDATE People SET areaID=?, typeID=?, fName=?, lName=?, phone=?, address=?, comments=?, latitude=?, longitude=? WHERE id=?""", (
                area,
                group,
                fname,
                lname,
                phone,
                address,
                comments,
                lat,
                lng,
                ID
            ))

            self.save()

    def updatePeopleAreas( self ):
        people = self.getPeople()

        if people:

            for person in people:

                area = self.getAreaFromPoint( person[8], person[9] )

                if area[0] != person[1]:

                    self.db.execute("""UPDATE People SET areaID=? WHERE id=?""", (
                        area[0],
                        person[0]
                    ))

    # ==================================================================
    # Extra
    # ==================================================================
    def pointInPoly( self, x, y, poly ):
        # Crazyness that I didn't make.

        # check if point is a vertex
        if (x,y) in poly: return True

        # check if point is on a boundary
        for i in range(len(poly)):
            p1 = None
            p2 = None
            if i==0:
                p1 = poly[0]
                p2 = poly[1]
            else:
                p1 = poly[i-1]
                p2 = poly[i]
            if p1[1] == p2[1] and p1[1] == y and x > min(p1[0], p2[0]) and x < max(p1[0], p2[0]):
                return True

        n = len(poly)
        inside = False

        p1x,p1y = poly[0]
        for i in range(n+1):
            p2x,p2y = poly[i % n]
            if y > min(p1y,p2y):
                if y <= max(p1y,p2y):
                    if x <= max(p1x,p2x):
                        if p1y != p2y:
                            xints = (y-p1y)*(p2x-p1x)/(p2y-p1y)+p1x
                        if p1x == p2x or x <= xints:
                            inside = not inside
            p1x,p1y = p2x,p2y

        if inside:
            return True

        return False

    def save( self ):
        self.connection.commit()

    def close( self ):
        self.save()
        self.connection.close()
