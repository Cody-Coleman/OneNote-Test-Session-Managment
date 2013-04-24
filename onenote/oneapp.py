'''
Created on Nov 10, 2011

@author: ccoleman
'''

import win32com.client
import re
from session2xml.session2xml import strip_span
from xml.etree import ElementTree

if win32com.client.gencache.is_readonly == True:

    #allow gencache to create the cached wrapper objects
    win32com.client.gencache.is_readonly = False

    # under p2exe the call in gencache to __init__() does not happen
    # so we use Rebuild() to force the creation of the gen_py folder
    win32com.client.gencache.Rebuild()

    # NB You must ensure that the python...\win32com.client.gen_py dir does not exist
    # to allow creation of the cache in %temp%


onapp = win32com.client.gencache.EnsureDispatch( 'OneNote.Application' )


def getSimpleHierarchy():
    return ElementTree.fromstring( onapp.GetHierarchy( "", win32com.client.constants.hsPages ) )

# Returns the Notebook Hierarchy as a Dictionary Array  
def getHierarchy( NS ):
    tmpstr = onapp.GetHierarchy( "", win32com.client.constants.hsPages )
    tmpstr = tmpstr.encode( 'ascii', 'ignore' )
    oneTree = ElementTree.fromstring( tmpstr )

    notebooks = []

    for notebook in oneTree:
        nbk = parseAttributes( notebook )
        if ( notebook.getchildren() ):
            s, sg = getSections( notebook, NS )
            if ( s != [] ):
                nbk['sections'] = s

            # Removes RecycleBin from SectionGroups and adds as a first class object
#            for i in range( len( sg ) ):
#                if ( 'isRecycleBin' in sg[i] ):
#                    nbk['recycleBin'] = sg[i]
#                    sg.pop( i )

            if ( sg != [] ):
                nbk['sectionGroups'] = sg

        notebooks.append( nbk )

    return notebooks

# Takes in a Notebook or SectionGroup  and returns a Dict Array of its Sections & Section Groups
def getSections( notebook, NS ):
    sections = []
    sectionGroups = []
    for section in notebook:
        if ( section.tag == NS + "SectionGroup" ):
            newSectionGroup = parseAttributes( section )
            if ( section.getchildren() ):
                s, sg = getSections( section, NS )
                if ( sg != [] ):
                    newSectionGroup['sectionGroups'] = sg
                if ( s != [] ):
                    newSectionGroup['sections'] = s
            sectionGroups.append( newSectionGroup )

        if ( section.tag == NS + "Section" ):
            newSection = parseAttributes( section )
            if ( section.getchildren() ):
                newSection['pages'] = getPages( section )
            sections.append( newSection )

    return sections, sectionGroups

# Takes in a Section and returns a Dict Array of its Pages
def getPages( section ):
    pages = []
    for page in section:
        newPage = parseAttributes( page )
        if ( page.getchildren() ):
            newPage['meta'] = getMeta( page )
        pages.append( newPage )
    return pages

# Takes in a Page and returns a Dict Array of its Meta properties
def getMeta ( page ):
    metas = []
    for meta in page:
        metas.append( parseAttributes( meta ) )
    return metas

# Takes in an object and returns a dictionary of its values
def parseAttributes( obj ):
        tempDict = {}
        for key, value in obj.items():
            tempDict[key] = value
        return tempDict

# Returns the xml tree structure of the page content    
def getPageContent( pgid, NS, verbose ):
    try:
        tmpstr = onapp.GetPageContent( pgid )
    except:
        print "Failed to parse contents of page"
    else:
        # NORMALIZE TMPSTR INCASE THERE ARE THINGS THAT ELEMENTREE CHOKES ON
        tmpstr = tmpstr.encode( 'ascii', 'ignore' )
        span = re.compile( r"<span.*?>", re.IGNORECASE | re.DOTALL )
        endspan = re.compile( r'</span>' )
        schema = 'one:Page xmlns:one="http://schemas.microsoft.com/office/onenote/2010/onenote"'
        one_tag = 'one:'
        tmpstr = re.sub( span, '', tmpstr )
        tmpstr = re.sub( endspan, '', tmpstr )
        tmpstr = re.sub( schema, 'Page', tmpstr )
        tmpstr = re.sub( one_tag, '', tmpstr )
        if verbose:
            tmpfile = open( 'xml_output_from_onenote.xml', 'w' )
            tmpfile.write( tmpstr )
            tmpfile.close()
        oneTree = ElementTree.fromstring( tmpstr )
        return oneTree



def getID( book, pName, pgid='' ):
    for elem in book:
        if 'pages' in elem:
            for p in elem['pages']:
                if ( p['name'].lower().find( pName ) >= 0 ):
                    pgid = p['ID']
                    return pgid
        if 'sectionGroups' in elem:
            pgid = getID( elem['sectionGroups'], pName, pgid )
            if len( pgid ) > 0:
                return pgid
        if 'sections' in elem:
            pgid = getID( elem['sections'], pName, pgid )
            if len( pgid ) > 0:
                return pgid
    return pgid


def getT( page, NS ):
    from cStringIO import StringIO
    fname = StringIO()
    findT( page, fname, NS )
    return fname.getvalue()


# writes to a file the output T's found in the page should still look into making this just a string object in memory
def findT( book, filename, NS, cell=False ):

    for i in range( len( book ) ):
        if book[i].tag == NS + 'Outline':
            filename.write( '\nOUTLINE:' )
        elif book[i].tag == NS + 'Table':
            filename.write( "\nTABLE" )
        elif book[i].tag == NS + 'Row':
            filename.write( '\n|-\n' )
        elif book[i].tag == NS + 'Cell':
            cell = True

        if( book[i].getchildren() ):
            findT( book[i], filename, NS, cell )
        elif( book[i].tag == NS + 'T' ):
            #if book[i].text != None:
            if cell:
                filename.write( '||  ' )#%s ' % book[i].text.strip())
                if book[i].text:
                    filename.write( book[i].text.strip() + '  ' )
                else:
                    filename.write( ' ' )
                cell = False
            else:
                if book[i].text:
                    filename.write( book[i].text.strip() + '\n' )
                else:
                    filename.write( ' ' )


def getOutline( page, NS ):
    outline = []
    for p in page:
        if p.tag == '%sOutline' % NS:
            outline.append( p )
    return outline