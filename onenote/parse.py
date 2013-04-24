'''
Created on Nov 30, 2011

@author: ccoleman
'''
from xml.etree import ElementTree
from session2xml.session2xml import strip_span


#NS = "{http://schemas.microsoft.com/office/onenote/2007/onenote}"
#NS = "{http://schemas.microsoft.com/office/onenote/%s/onenote}"

def get_title( page ):
    tmp = page.find( './/T' )
    title_text = strip_span( tmp.text )
    return title_text.replace( '<a\n', '<a ' )

def get_categories( page ):
    outlines = page.findall( './/Outline' )
    cat_type = {}
    for out in outlines:
        tmp = out.find( './/T' )
        if tmp.text is not None:
            tmp_text = strip_span( tmp.text ).strip()
            cat_type[( tmp_text.replace( ':', '' ) ).lower()] = out
    return cat_type

def get_testers( outline ):
    tmp = outline.findall( './/T' )
    tmparray = []
    # FIRST INDEX IS ALWAYS THE TITLE OF THE OUTLINE
    for i in range( 1, len( tmp ) ):
        if tmp[i].text is not None:
            tmparray.append( tmp[i].text )
    return tmparray

def get_table( outline ):
    tbl_array = []
    tbl = outline.find( './/Table' )
    rows = tbl.findall( './/Row' )
    # First value in rows is alwasy the header
    for i in range( 1, len( rows ) ):
        cell_array = []
        # GET THE CELLS THEN FIGURE OUT IF THERE ARE MULTIPLE T'S IN A CELL
        cells = rows[i].findall( './/Cell' )
        for j in range( 0, len( cells ) ):
            tmp_text = ''
            tees = cells[j].findall( './/T' )
            # MULTIPLE TEES NEED TO BE SMASHED TOGETHER
            for k in range( 0, len( tees ) ):
                if tees[k].text is not None:
                    tmp = tees[k].text.strip()
                   # if len( tmp ) > 0:
                    tmp_text = ''.join( ( tmp_text, ' ', tees[k].text ) ).strip()
                    #else:
                    #    tmp_text = None
            if len( tmp_text ) <= 0:
                tmp_text = None
            cell_array.append( tmp_text )
        for i in range( 0, len( cell_array ) ):
            if cell_array[i] is not None:
                cell_array[i] = cell_array[i].strip()
        tbl_array.append( cell_array )
    return filter( clean_table, tbl_array )

def clean_table( array ):
    keep = True
    for row in array:
            if row == None:
                keep = False
            else:
                keep = True
                break
    return keep
