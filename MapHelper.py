# 1.0
import json
import os
import string

from itertools import chain

MAP_TYPE_NORMAL        = 1
MAP_TYPE_REGION_HELPER = 2

class MapHelper:
    def __init__( self, name, wbMap, ptype=1 ):
        self._name = name
        self._map  = wbMap
        self._type = ptype

        self.buildTemplate()
        self.createTempFile()

        # what the nasty thing makes
        self.peopleCharacters = ('0', '1', '2', '3', '4', '5', '6', '7', '8', '9', 'A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T', 'U', 'V', 'W', 'X', 'Y', 'Z', 'AA', 'BB', 'CC', 'DD', 'EE', 'FF', 'GG', 'HH', 'II', 'JJ', 'KK', 'LL', 'MM', 'NN', 'OO', 'PP', 'QQ', 'RR', 'SS', 'TT', 'UU', 'VV', 'WW', 'XX', 'YY', 'ZZ')

    def createTempFile( self ):
        self.filePath = '%s/Map%s.html' % ( os.path.dirname(os.path.abspath(__file__)), self._name )
        self.file = open(self.filePath, 'w+b')

    def cleanup( self ):
        # Clean up the temporary file yourself
        self.file.close()
        os.remove(self.filePath)

    def buildPage( self, *args, **kwargs ):
        if self._type == MAP_TYPE_NORMAL:
            self.buildNormalPage( *args, **kwargs )

        elif self._type == MAP_TYPE_REGION_HELPER:
            self.buildRegionHelperPage( *args, **kwargs )

    def buildNormalPage( self, areas=None, people=None ):
        sAreas = ''

        if areas:
            if type(areas[0]) in ( list, tuple ):
                for area in areas:
                    sAreas += '            ["{}","{}",{},"{}",{},{}],\n'.format( area[1], area[2], area[3], area[4], area[5], json.loads(area[6]) )
            else:
                sAreas += '            ["{}","{}",{},"{}",{},{}],\n'.format( areas[1], areas[2], areas[3], areas[4], areas[5], json.loads(areas[6]) )

        sPeople = ''

        if people:
            if type(people[0]) in ( list, tuple ):
                for person in people:
                    sPeople += '            ["{}","{}","{}","{}","{}","{}",{},{}],\n'.format(person[10], person[3], person[4], person[5], person[6], person[7], person[8], person[9])
            else:
                sPeople += '            ["{}","{}","{}","{}","{}","{}",{},{}],\n'.format(people[10], people[3], people[4], people[5], people[6], people[7], people[8], people[9])

        try:
            self.file.seek(0)
            self.file.truncate()
            self.file.write( self._template( polygonInfo = sAreas, peopleInfo = sPeople ) )
            self.file.flush()

            self._map.Navigate( "File://%s" % self.filePath )
        except Exception as e:
            self.cleanup()

            raise e

    def buildRegionHelperPage( self, fColor="#999999", fOpacity=0.1, sColor="#999999", sOpacity=0.7, coordinates='', editable=True, clat = 33.986578, clng = -118.169661 ):
        # So you can pass in the DB row entry for the area
        if type( fColor ) in ( list, tuple ):
            coordinates = json.loads( fColor[6] )
            sOpacity    = fColor[3]
            sColor      = fColor[2]
            fOpacity    = fColor[5]
            fColor      = fColor[4]

        s = ''
        if type(coordinates) in ( list, tuple ):
            for coord in coordinates:
                s += '{}, {}\n'.format( coord[0], coord[1] )

        try:
            self.file.seek(0)
            self.file.truncate()
            self.file.write( self._template( fillColor=fColor, fillOpacity=fOpacity, strokeColor=sColor, strokeOpacity=sOpacity, coordinates=s, editable=str(editable).lower(), cLat = clat, cLng = clng ) )
            self.file.flush()

            self._map.Navigate( "File://%s" % self.filePath )
        except Exception as e:
            self.cleanup()

            raise e

    def getRegionPointsAsPeople( self, region ):
        points = json.loads(region[6])
        r = list()

        for i in xrange(len(points)): #ascii_uppercase
            r.append( ( '', '', '', '', 'Latitude: ' + str(points[i][0]), '', 'Longitude: ' + str(points[i][1]), '', points[i][0], points[i][1], self.peopleCharacters[i] ) )

        return r

    def buildTemplate( self ):
        if self._type == MAP_TYPE_NORMAL:
            self._template = """
<!DOCTYPE html>
<html>
<head>
    <!-- This crap is needed for IE -->
    <meta http-equiv="X-UA-Compatible" content="IE=edge" />
    <style type="text/css">
        html, body, #map-canvas {{ height: 100%; margin: 0; padding: 0;}}
    </style>

    <script type="text/javascript" src="https://maps.googleapis.com/maps/api/js?key=AIzaSyAhJoJcvbNNGH_Sfn4UJRNZU6CZCDpEq3I"></script>
    <script src="https://ajax.googleapis.com/ajax/libs/jquery/2.1.3/jquery.min.js"></script>

    <script type="text/javascript">
        var polygonInfo = [
{polygonInfo}
        ];

        var peopleInfo = [
{peopleInfo}
        ];

        var markers = [];
        var polygons = {{}};
        var infoWindow;
        var map;

        function initialize() {{

            var center = new google.maps.LatLng(33.986578, -118.169661);


            if ( polygonInfo.length > 0 ) {{
                center = getCenterPoint( polygonInfo );

                center = new google.maps.LatLng( center[0], center[1] );
            }}

            var mapOptions = {{
                zoom: 14,
                center: center,
                mapTypeId: google.maps.MapTypeId.ROADMAP // Roadmap so you can zoom more!
            }};

            map        = new google.maps.Map(document.getElementById('map-canvas'), mapOptions);
            infoWindow = new google.maps.InfoWindow();

            drawPolygons();
            drawPeople();
        }}

        function drawPolygons() {{
            for ( area in polygonInfo ) {{
                area = polygonInfo[area];

                var paths = [];
                for ( point in area[5] ) {{
                    point = area[5][point];
                    paths.push( new google.maps.LatLng( point[0], point[1] ) );
                }}

                polygons[area[0]] = new google.maps.Polygon({{
                    strokeColor  : area[1],
                    strokeOpacity: area[2],
                    fillColor    : area[3],
                    fillOpacity  : area[4],
                    paths        : paths
                }});

                polygons[area[0]].setMap( map );

                google.maps.event.addListener( polygons[area[0]], 'click', function(event){{
                    for ( var area in polygons ){{
                        if ( polygons[area] == this ){{
                            infoWindow.setContent( area );
                            infoWindow.setPosition( event.latLng );
                            infoWindow.open(map);
                            break;
                        }}
                    }}
                }});
            }}
        }}

        function drawPeople() {{
            for ( person in peopleInfo ) {{
                person = peopleInfo[person];

                markers.push( new google.maps.Marker({{
                    icon: "http://www.googlemapsmarkers.com/v1/"+person[0]+"/FF6961/",
                    position: new google.maps.LatLng( person[6], person[7] ),
                    map: map,
                    title: person[2] + ", " + person[1] + "\\n" + person[4] + "\\n" + person[3]
                }}));
            }}
        }}

        function getRegionMinMax( points ) {{
            var xSmall = parseFloat(points[0][0]);
            var xLarge = parseFloat(points[0][0]);
            var ySmall = parseFloat(points[0][1]);
            var yLarge = parseFloat(points[0][1]);

            var x, y;
            for ( point in points ) {{
                x = parseFloat(points[point][0]);
                y = parseFloat(points[point][1]);

                if ( x < xSmall ) {{ xSmall = x; }}
                if ( x > xLarge ) {{ xLarge = x; }}

                if ( y < ySmall ) {{ ySmall = y; }}
                if ( y > yLarge ) {{ yLarge = y; }}
            }}

            return [ [ xSmall, ySmall ], [ xLarge, yLarge ] ];
        }}

        function getCenterPoint( points ) {{
            allPoints = [];

            for (area in points) {{
                area = polygonInfo[area];

                for ( point in area[5] ) {{
                    point = area[5][point];
                    allPoints.push( point );
                }}
            }}

            if (allPoints.length == 0) {{
                return [ 33.986578, -118.169661 ];
            }}


            var minMax = getRegionMinMax( allPoints );

            var min = minMax[0];
            var max = minMax[1];

            lat = min[0] + ((max[0] - min[0]) / 2);
            lng = min[1] + ((max[1] - min[1]) / 2);

            return [ lat, lng ];
        }}

        // Initialize the map
        google.maps.event.addDomListener(window, 'load', initialize);

    </script>

</head>

<body>
    <div id="map-canvas"></div>
</body>

</html>""".format

        if self._type == MAP_TYPE_REGION_HELPER:
            self._template = """
<!DOCTYPE html>
<html>
    <head>
        <!-- This crap is needed for IE -->
        <meta http-equiv="X-UA-Compatible" content="IE=edge" />
        <style type="text/css">
            html, body, #map-canvas {{ height: 100%; margin: 0; padding: 0;}}
            #polygon-helper {{
                position: absolute;
                bottom: 25px;
                left: 10px;
            }}

            #coordinates {{
                height: 200px;
                width: 200px;
            }}
        </style>

        <script type="text/javascript" src="https://maps.googleapis.com/maps/api/js?key=AIzaSyAhJoJcvbNNGH_Sfn4UJRNZU6CZCDpEq3I"></script>
        <script src="https://ajax.googleapis.com/ajax/libs/jquery/2.1.3/jquery.min.js"></script>

        <script id="regionColors" type="text/javascript">
            var fColor   = '{fillColor}';
            var fOpacity = {fillOpacity};
            var sColor   = '{strokeColor}';
            var sOpacity = {strokeOpacity};
        </script>

        <script type="text/javascript">

            var map;
            var polygon;

            function initialize() {{

                //var center = new google.maps.LatLng(33.986578, -118.169661);
                var center = new google.maps.LatLng({cLat}, {cLng});

                if ( $.trim($('#coordinates').val()) != '' ) {{
                    var center = getCenterPoint( getPointsFromText( false ) );

                    center = new google.maps.LatLng( center[0], center[1] );
                }}

                var mapOptions = {{
                    zoom: 15,
                    center: center,
                    mapTypeId: google.maps.MapTypeId.ROADMAP // Roadmap so you can zoom more!
                }};

                map = new google.maps.Map(document.getElementById('map-canvas'), mapOptions);

                google.maps.event.addListener( map, "click", function(event) {{
                    var points;
                    var lat = event.latLng.lat();
                    var lng = event.latLng.lng();

                    // if there is no polygon, create it
                    if ( polygon == undefined ) {{
                        points = [ new google.maps.LatLng( lat, lng ) ];
                        drawPolygon( points );

                    // if there is a polygon add the new point to it
                    }} else {{
                        points = getPolygonPoints( true );
                        // get rid of the old polygon
                        polygon.setPath([]);
                        points.push( new google.maps.LatLng( lat, lng ) );
                        drawPolygon( points );
                    }}

                    updateText();

                }});

                if ( $.trim($('#coordinates').val()) != '' ) {{
                    drawPolygon( getPointsFromText(true) );
                }}

                $('#coordinates').keyup( function( data ) {{
                    if ( data.keyCode == 86 ) {{ return null; }}
                    // if ( polygon != undefined ) {{ polygon.setPath([]); }}
                    var points = getPointsFromText(true);
                    if (points.length > 0) {{
                        drawPolygon( getPointsFromText(true) );
                        map.setCenter( points[0] );
                    }} else {{
                        if ( polygon != undefined ) {{ polygon.setPath([]); }}
                    }}
                }});

                $('#coordinates').toggle({editable});
            }}

            function getRegionMinMax( points ) {{
                var xSmall = parseFloat(points[0][0]);
                var xLarge = parseFloat(points[0][0]);
                var ySmall = parseFloat(points[0][1]);
                var yLarge = parseFloat(points[0][1]);

                var x, y;
                for ( point in points ) {{
                    x = parseFloat(points[point][0]);
                    y = parseFloat(points[point][1]);

                    if ( x < xSmall ) {{ xSmall = x; }}
                    if ( x > xLarge ) {{ xLarge = x; }}

                    if ( y < ySmall ) {{ ySmall = y; }}
                    if ( y > yLarge ) {{ yLarge = y; }}
                }}

                return [ [ xSmall, ySmall ], [ xLarge, yLarge ] ];
            }}

            function getCenterPoint( points ) {{
                var minMax = getRegionMinMax( points );

                var min = minMax[0];
                var max = minMax[1];

                lat = min[0] + ((max[0] - min[0]) / 2);
                lng = min[1] + ((max[1] - min[1]) / 2);

                return [ lat, lng ];
            }}

            function getPointsFromText( googlePoints ) {{
                // Default paramater if you don't pass true, it will return just x y values
                googlePoints = googlePoints || false;

                var points = [];

                var re = /(.*), (.*)/

                var lines = $.trim($('#coordinates').val()).split("\\n");

                for ( line in lines ) {{
                    line = $.trim(line);

                    if (lines[line] != "" ) {{

                        info = lines[line].match(re);

                        if ( info != null ) {{
                            // console.log( info[1], info[2] );
                            if (googlePoints == true) {{
                                points.push( new google.maps.LatLng( info[1], info[2] ) );
                            }} else {{
                                points.push( [info[1], info[2]] );
                            }}
                        }}
                    }}
                }}

                return points;
            }}

            function updateInfo( fcolor, fopacity, scolor, sopacity ) {{
                fColor   = fcolor;
                fOpacity = fopacity;
                sColor   = scolor;
                sOpacity = sopacity;
            }}

            function drawPolygon( points ) {{
                points = points || getPolygonPoints(true);

                if ( points == null || points == undefined || points.length < 1 ) {{ return; }}
                if ( polygon != undefined ) {{ polygon.setPath([]); }}

                polygon = new google.maps.Polygon({{
                    paths         : points,
                    strokeColor   : sColor,
                    fillColor     : fColor,
                    strokeWeight  : 2,
                    strokeOpacity : sOpacity,
                    fillOpacity   : fOpacity,
                    editable      : {editable}
                }});

                google.maps.event.addListener( polygon.getPath(),    'set_at', function() {{ updateText(); }} );
                google.maps.event.addListener( polygon.getPath(), 'insert_at', function() {{ updateText(); }} );

                polygon.setMap(map);

                return 'drawn';
            }}

            function getPolygonPoints( googlePoints ) {{
                // Default paramater if you don't pass true, it will return just x y values
                googlePoints = googlePoints || false;

                if ( polygon == undefined ) {{ return []; }}

                vertices = polygon.getPath();

                ret = [];

                for ( var i = 0; i < vertices.getLength(); i++ ) {{
                    var xy = vertices.getAt(i);

                    if (googlePoints == true) {{
                        ret.push( xy ); //google.maps.LatLng( xy.lat(), xy.lng() ) );
                    }} else {{
                        ret.push( [xy.lat(), xy.lng()] );
                    }}
                }}

                return ret;
            }}

            function updateText( polyPoints ) {{
                polyPoints = getPolygonPoints();

                str = "";
                for ( point in polyPoints ) {{
                    str += parseFloat(polyPoints[point][0]).toFixed(6) + ", " + parseFloat(polyPoints[point][1]).toFixed(6) + "\\n";
                }}

                $("#coordinates").val( str );
            }}

            function SetCenter( latitude, longitude ) {{
                alert( 'set center: ' + latitude, longitude );
                map.setCenter( new google.maps.LatLng( latitude, longitude ) );
            }}

            google.maps.event.addDomListener(window, 'load', initialize);
        </script>

    </head>
    <body>
        <div id="map-canvas"></div>
        <div id="polygon-helper">
            <form>
                <textarea id="coordinates">{coordinates}</textarea>
            </form>
        </div>
    </body>
</html>
            """.format
