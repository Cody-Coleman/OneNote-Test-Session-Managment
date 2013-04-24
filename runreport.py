'''
Created on Nov 10, 2011

@author: ccoleman
'''
from onenote.oneapp import getHierarchy, getSimpleHierarchy, getID
from onenote.oneapp import getPageContent, getT
from onenote.parse import get_categories, get_title, get_testers, get_table
from session2xml import session2xml as s2x
from session2xml import writexml as wxml
import argparse
import sys
from xml.etree import ElementTree as ET

prompt = False
team = None
release = None
NS = "{http://schemas.microsoft.com/office/onenote/%s/onenote}"


def get_args():
    parser = argparse.ArgumentParser( description='''Parse out a onenote file 
                                                    and generate an xml then 
                                                    upload it to a server for 
                                                    viewing''' )

    parser.add_argument( '-p', '--path', default='',
                        help='The path for xml output' )

    parser.add_argument( '-n', '--name',
                        help='The name of the onenote page - REQUIRED' )

    parser.add_argument( '-t', '--team',
                        help='''The name of the product team this report is 
                        for - example Mobile, PC, MAC, FeatureGroup - defaults to 
                        "team"''' )

    parser.add_argument( '-r', '--release',
                        help='''The name of the release this report is for - 
                        example 1.3, 2.0, bantha, buzz - 
                        defaults to "realease"''' )

    parser.add_argument( '-u', '--url',
                        default='http://localhost/uploads/uploader.php',
                        help='Path to the server on where to upload' )

    parser.add_argument( '-s', '--section', default='',
                        help='''Path of Section to upload, mutually exclusive 
                        to page - Usage {Workbook}/{sectionGroup}*/{section} 
                        where you can have more than one nested section''' )

    parser.add_argument( '-o', '--onenote', default='2007',
                        help='''The version of onenote such as 2007 
                        or 2010''' )

    parser.add_argument( '-v', '--verbose', action='store_true',
                        help='write out the actions taken by the script' )

    return parser.parse_args()

def upload_file( xmlfile, path, team='', release='' ):
    import MultipartPostHandler, urllib2
    opener = urllib2.build_opener( MultipartPostHandler.MultipartPostHandler )
    params = {'team': team, 'release': release, 'report':open( xmlfile, 'rb' )}
    try:
        ret = opener.open( path, params ).read()
    except:
        print "failed to connect to server %s" % path
        sys.exit()
    else:
        return ret


def parse2xml( args, ID='' ):
    ############
    # VARIABLES
    ############
    inum = 0
    onum = 0
    bnum = 0
    b_time = 0
    t_time = 0
    s_time = 0
    session_length = 90
    session_actual = 90
    global team
    global release
    global sprint
    global prompt

    root = ET.Element( 'session' )
    if args.verbose:
        print "ID: %s" % ID
    page = getPageContent( ID, NS, args.verbose )
    if args.verbose:
        print "Finding the outlines to the page"
    content = get_categories( page )
    if args.verbose:
        print content.keys()
        tree = ET.ElementTree( page )
        wxml.indent( tree.getroot() )
        tree.write( "page.xml" )
    #===========================================================================
    # CHARTER
    #===========================================================================
    title = get_title( page )
    charter = s2x.split_a( title )
    charter['title'] = ( s2x.strip_special( charter['title'] ) ).strip()
    charter['text'] = ( s2x.strip_special( charter['text'] ) ).strip()
    fpath = args.path + "_" + charter['title'] + ' ' + charter['text'] + '.xml'
    if args.verbose:
        print "opening file %s to store onenote XML generation" % fpath
    ET.SubElement( root, 'name' ).text = charter['text']
    if charter['title']:
        ET.SubElement( root, 'bli' ).text = charter['title']
    if charter['link']:
        ET.SubElement( root, 'bli_link' ).text = charter['link']
    if args.verbose:
        print "charter: %s" % charter
    status = s2x.get_status( title )
    ET.SubElement( root, 'status' ).text = status
    if args.verbose:
        print "Status: %s" % status
    #===========================================================================
    # SESSION DATA
    #===========================================================================
    if 'session data' in content:
        session_data = get_table( content['session data'] )
        if len( session_data ) > 0:
            if session_data[0][0]:
                team = session_data[0][0].lower().strip()
            if session_data[0][1]:
                release = session_data[0][1].lower().strip()
            if session_data[0][2]:
                ET.SubElement( root, 'sprint' ).text = session_data[0][2]
            if session_data[0][3]:
                tmp = s2x.split_a( session_data[0][3] )
                ET.SubElement( root, 'bli' ).text = tmp['title']
                ET.SubElement( root, 'bli_link' ).text = tmp['link']
    while( not team ):
        prompt = True
        team = raw_input( 'Name of Team: ' )
    while( not release ):
        prompt = True
        release = raw_input( 'Name of Release: ' )
    #===========================================================================
    # TESTERS
    #===========================================================================
    if 'testers' in content:
        tst_array = get_testers( content['testers'] )
        testers = ET.SubElement( root, 'testers' )
        for val in tst_array:
            ET.SubElement( testers, 'tester_name' ).text = val
            if args.verbose:
                print "testers: %s" % val
    #===========================================================================
    # SESSION SIZE
    #===========================================================================
    if 'session size' in content:
        #=======================================================================
        # REMINDER: GET_TABLE REUTRNS AN ARRAY OF ARRAYS FOR EACH CELL
        # SO SESSION[0] = THE FIRST ROW AND SESSION[0][0] IS THE FIRST CELL
        # OF THAT ROW, SESSION[0][1] IS THE SECOND CELL OF THAT ROW
        #=======================================================================
        session = get_table( content['session size'] )
        start_t = "12/1/2011, 10:00 AM"
        end_t = "12/1/2011, 11:30 AM"
        if 'activities' in content:
            activities = get_table( content['activities'] )
            if len( activities ) > 0:
                start_t = activities[0][0]
                counter = 1
                while True:
                    end_t = activities[len( activities ) - counter][0]
                    if end_t is not None:
                        break
                    else:
                        counter += 1
        session_length, session_type = s2x.find_session_length( session[0], start_t, end_t )
        ET.SubElement( root, 'session_type' ).text = str( session_type )
        ET.SubElement( root, 'session_length' ).text = str( session_length )
        if args.verbose:
            print "session size in minutes: %d" % session_length

    if 'activities' in content:
        activities = get_table( content['activities'] )
        if 'session size' in content:
            if session_length:
                t_time = s2x.get_time( activities, 'T', session_length )
                b_time = s2x.get_time( activities, 'B', session_length )
                s_time = s2x.get_time( activities, 'S', session_length )
            else:
                t_time = s2x.get_time( activities, 'T', 90 )
                b_time = s2x.get_time( activities, 'B', 90 )
                s_time = s2x.get_time( activities, 'S', 90 )
            if args.verbose:
                print ( "sTime: %s\ntTime: %s\nbTime: %s" %
                      ( s_time, t_time, b_time ) )
            session_time = ET.SubElement( root, 'session_time' )
            ET.SubElement( session_time, 'b_time' ).text = str( b_time )
            ET.SubElement( session_time, 't_time' ).text = str( t_time )
            ET.SubElement( session_time, 's_time' ).text = str( s_time )
    #===========================================================================
    # DEFECTS      
    #===========================================================================
    if 'defects' in content:
        defects = get_table( content['defects'] )
        defects = s2x.find_defects( defects )
        bnum = len( defects )
        for val in defects:
            defect_node = ET.SubElement( root, 'defect' )
            ET.SubElement( defect_node, 'defect_title' ).text = val['text']
            ET.SubElement( defect_node, 'defect_number' ).text = val['title']
            ET.SubElement( defect_node, 'defect_link' ).text = val['link']
            if args.verbose:
                print "defect: %s" % val
    #===========================================================================
    # ACTIVITIES
    #===========================================================================
    if 'activities' in content:
        activities = get_table( content['activities'] )
        for val in activities:
            if val[2] == 'I':
                inum += 1
            elif val[2] == 'O':
                onum += 1
        # If we got issues or opportunies, write these
        parent_dict = {
                     'O':ET.SubElement( root, 'opportunity' ),
                     'I':ET.SubElement( root, 'issue' ),
                     'T':ET.SubElement( root, 'test' ),
                     'S':ET.SubElement( root, 'setup' ),
                     'B':ET.SubElement( root, 'bug' )
                     }
        node_dict = {
                       'O':'o_node',
                       'I':'i_node',
                       'T':'t_node',
                       'S':'s_node',
                       'B':'b_node'
                       }
        for val in activities:
            if val[2] in parent_dict:
                ET.SubElement( 
                              parent_dict[val[2]],
                              node_dict[val[2]]
                              ).text = val[1]
                if args.verbose:
                    print "Type: %s, Text: %s" % ( val[2], val[1] )
    #===========================================================================
    # COMPLETE XML AND UPLOAD RESULTS
    #===========================================================================
    session_count = ET.SubElement( root, 'session_count' )
    ET.SubElement( session_count, 'b_count' ).text = str( bnum )
    ET.SubElement( session_count, 'o_count' ).text = str( onum )
    ET.SubElement( session_count, 'i_count' ).text = str( inum )
    if args.verbose:
        print "Issue Count: %d\nOpportunities Count: %d" % ( inum, onum )
    print ( "sTime: %s\ntTime: %s\nbTime: %s" %
          ( s_time, t_time, b_time ) )
    print ( "Actual Session Time: %d" % session_length )
    print "iCount: %s\nbCount: %s\noCount: %s" % ( inum, bnum, onum )
    tree = ET.ElementTree( root )
    wxml.indent( tree.getroot() )
    tree.write( fpath )
    ret = upload_file( fpath, args.url, team, release )
    if 'uploaded' in ret:
        print 'report uploaded succesfully'
    else:
        print ret
    if args.verbose:
        print ret + "\n"

if __name__ == '__main__':
    args = get_args()
    #===========================================================================
    # NEED TO FIND OUT IF THE REPORT IS ON A PAGE OR AN ENTIRE SECTION
    # IF NAME IS SET, IGNORE SECTION AND ASSUME NAME
    #===========================================================================
    team = args.team
    release = args.release
    pName = args.name
    if not args.section:
        while( not pName ):
            prompt = True
            pName = raw_input( 'Name of Charter: ' )
        # lower everything to make it easier
        pName = pName.lower()
        if args.verbose:
            print "gathering onenote file structure"
        NS = NS % args.onenote
        book = getHierarchy( NS )
        if args.verbose:
            print "finding ID for page %s" % pName
        pgid = getID( book, pName )

        if args.verbose:
            print "getting page contents and parsing, this can take some time"
        if len( pgid ) > 0:
            parse2xml( args, pgid )
        else:
            print "did not find page by that name\nexiting program"
            sys.exit()

    else:
        if args.verbose:
            print "Ready for sections %s" % args.section
        # PARSE THE SECTION STRING PASSED INTO EXPECTED VALUES
        path = args.section.split( '/' )
        # THE FIRST ELEMENT IS THE WORKBOOK AND THE LAST IS THE SECTION
        workbook = path[0].lower()
        section = path[len( path ) - 1].lower()
        # EVERY THING IN THE MIDDLE, IF IT EXISTS, IS THE SECTION GROUPS
        sectionGroup = path[1:len( path ) - 1]
        # NOW GET THE HIEARCHY INTO AN XML OBJECT TO USE TO FIND ALL THE IDS
        xmltree = getSimpleHierarchy()
        # TRAVERSE THE TREE AND FIND THE WORKBOOK
        found_workbook = False
        for child in xmltree:
            if 'name' in child.keys():
                if child.attrib['name'].lower() == workbook:
                    xmltree = child
                    found_workbook = True
                    break
        if not found_workbook:
            print "Workbook %s was not found in onenote" % workbook
            sys.exit()
        #xmltree.attrib['name']
        # NOW TRAVERSE THE TREE AGAIN (NOW THAT IT ONLY HAS THE ONE WORKBOOK)
        # AND FIND THE RIGHT SECTION GROUP
        found_section_group = False
        for i in range( len( sectionGroup ) ):
            for child in xmltree:
                if 'name' in child.keys():
                    if child.attrib['name'].lower() == sectionGroup[i].lower():
                        xmltree = child
                        found_section_group = True
                        break
        if len( sectionGroup ) > 0 and not found_section_group:
            print "Section groups %s not found" % sectionGroup
            sys.exit()
        # GET THE SECTION NAME AND ASSIGN THE CHILD TO THE TREE
        found_section = False
        for child in xmltree:
            if child.attrib['name'].lower() == section.lower():
                xmltree = child
                found_section = True
                break
        if not found_section:
            print "Section %s not found" % section
            sys.exit()
        print "Found %d pages to parse" % len( xmltree )
        for page in xmltree:
            parse2xml( args, page.attrib['ID'] )
        print "Sections uploaded"
    if prompt:
        raw_input( "report generated and uploaded, press enter to exit" )


