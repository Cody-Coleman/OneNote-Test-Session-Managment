'''
Created on Nov 11, 2011

@author: ccoleman
'''
from datetime import datetime as dt
import math
import re

def strip_span( some_string ):
    span = re.compile( r"<span.*?>", re.IGNORECASE | re.DOTALL )
    endspan = re.compile( r'</span>' )
    tmpstr = re.sub( span, '', some_string )
    tmpstr = re.sub( endspan, '', tmpstr )
    return tmpstr

def strip_dashes( some_string ):
    tmp = some_string.strip()
    if len( tmp ) > 0:
        if tmp[0] == '-':
            tmp = tmp[1:]
        if tmp[len( tmp ) - 1] == '-':
            tmp = tmp[:len( tmp ) - 1]
    return tmp.strip()

def strip_special( some_string ):
    op_pattern = re.compile( r'[^a-zA-Z0-9 ]' )
    tmp = re.sub( op_pattern, '', some_string )
    op_pattern = re.compile( r'  ' )
    return re.sub( op_pattern, ' ', tmp )

def get_status( charter ):
    status_list = {'*':'Not Started', '@':'Ready for Peer Review', '!':'Ready for Lead Review', '^':'In Progress'}
    tmp = charter.strip()
    char_val = tmp[0]
    if char_val in status_list:
        return status_list[char_val]
    else:
        return "Accepted"

def get_time( activity_array, activity_type, length=90 ):
    t_h = 0.00
    t_m = 0.00
    if len( activity_array ) > 0:
        t_start, t_end = find_time( activity_array, activity_type )

        for i in range( 0, len( t_start ) ):
            t_h += t_end[i].hour - t_start[i].hour
            t_m += t_end[i].minute - t_start[i].minute
    total_t = t_h * 60 + t_m
    return int( math.ceil( ( float( total_t ) / float( length ) ) * 100 ) )

def split_a( aref ):
    #sometimes Onenote puts a \n after the <a which is annoying
    # So going to strip it out and replace with a space if it exists
    aref = aref.replace( '\n', ' ' )

    start_of_link = '<a href="'
    end_of_link = '">'
    close_of_link = '</a>'

    if start_of_link in aref:
        link_str = aref[aref.find( start_of_link ) + len( start_of_link ):aref.find( end_of_link )].strip()
        title_str = aref[aref.find( end_of_link ) + len( end_of_link ):aref.find( close_of_link )].strip()
        text_str = aref[aref.find( close_of_link ) + len( close_of_link ):len( aref )].strip()
    else:
        link_str = ''
        title_str = ''
        text_str = aref
    return {'link': link_str, 'title': title_str, 'text': text_str}

def find_defects( defects ):
    defect_array = []
    op_pattern = re.compile( r'[a-zA-Z]*' )
    for defect in defects:
        # First cell is title, and link, second is text
        title = ''
        text = ''
        link = ''
        if defect[0] is not None:
            # strip out the link if it exists
            # split_a returns only a value in text if no link is found
            tmp_link = split_a( defect[0].lower().strip() )
            if len( tmp_link['title'] ) > 0:
                # found a linked title normalize text
                title = re.sub( op_pattern, '', tmp_link['title'].strip() )
                link = tmp_link['link']
            else:
                title = tmp_link['text']
        if defect[1] is not None:
            text = defect[1].strip()
        if len( text ) or len( title ) > 0:
                defect_array.append( {'title': title, 'link': link, 'text': text} )
    return defect_array

def find_time( activity_array, activity_type, format_style='%m/%d/%Y %I:%M %p' ):
    nutural_list = ['D', 'O', 'I']
    start_time = []
    end_time = []
    start_flag = False
    end_flag = False
    for i in range( 0, len( activity_array ) ):
        #if 'type' in activity_array[i] and activity_array[i]['type'] == activity_type:
        # time = 0, text = 1, type = 2 i = row
        if activity_array[i][2] == activity_type and activity_array[i][0] is not None:
            if not start_flag:
                start_flag = True
                end_flag = False
                start_time.append( dt.strptime( activity_array[i][0].replace( ',', '' ), format_style ) )
        elif activity_array[i][2] in nutural_list and activity_array[i][0] is not None:
            pass
        elif ( activity_array[i][2] != activity_type and
              activity_array[i][2] not in nutural_list and
              activity_array[i][0] is not None ):
            if not end_flag and ( end_flag != start_flag ):
                end_flag = True
                start_flag = False
                end_time.append( dt.strptime( activity_array[i][0].replace( ',', '' ), format_style ) )

    if len( end_time ) < len( start_time ):
        index = len( activity_array ) - 1
        while True:
            if activity_array[index][0] is not None:
                break
            else:
                index -= 1
        end_time.append( dt.strptime( activity_array[index][0].replace( ',', '' ), format_style ) )
    return start_time, end_time

def find_session_length( session, start_time, end_time, format_style='%m/%d/%Y %I:%M %p' ):
    session_dict = {'S' : 30, 'N': 90, 'L' : 120}
    s_type = ( session[0][0].upper() ).strip()
    start = dt.strptime( start_time.replace( ',', '' ), format_style )
    end = dt.strptime( end_time.replace( ',', '' ), format_style )
    t_h = end.hour - start.hour
    t_m = end.minute - start.minute
    s_length = t_h * 60 + t_m
    # USE THE +/- 15 TO DETERMINE CALCULATED TIME OF SESSION
    # RETURN THE TYPE AND LENGTH
    actual_type = 'Normal'
    if ( s_length <= 45 ):
        actual_type = 'Small'
    elif( s_length > 45 or s_length <= 105 ):
        actual_type = 'Normal'
    elif( s_length > 105 ):
        actual_type = 'Large'

    #low = session_dict[s_type] - 15
    #high = session_dict[s_type] + 15
#    if low <= s_length <= high:
#        return ( s_length, s_actual )
#    elif s_length < low:
#        return ( low, s_actual )
#    elif s_length > high:
#        return ( high, s_actual )
#    else:
#        return ( session_dict[s_type], session_dict[s_type] )
    return ( s_length, actual_type )