# Python PowerFactory API
try:
    import powerfactory as pf
    PF = True
except ModuleNotFoundError:
    print('Script Not Excecuted Using PowerFactory')
    PF = False

import sys
import numpy as np
import xml.dom.minidom
import pandas as pd
from bokeh.plotting import figure, show, output_file

# Ergon Energy Helper Functions for Logging
sys.path.append(r'\\Ecasd01\WksMgmt\PowerFactory\Scripts\pfTextOutputs')
import pftextoutputs
# Ergon Energy Helper Functions for lots of stuff, inc Ecorp ID extraction
sys.path.append(r'\\Ecasd01\WksMgmt\PowerFactory\ScriptsDEV\pfSharedFunctions')
import pfsharedfunctions as pfsf

#import CSV
import csv

# Logging, A more efficient way to print
import logging
logger = logging.getLogger(__name__)

# Ergon Energy Helper Functions for getting members
sys.path.append(
    r"\\Ecasd01\WksMgmt\PowerFactory\ScriptsDEV\ShortPFScripts\PrintScripts\PrintPFObjectMembers"
)
import printPFObjectV2_descs as printmembers
 
import datetime

import os

def run_main():
    app = pf.GetApplication()
    app.ClearOutputWindow()
    start_stuff(app)


def start_stuff(app, project=None):
    """
    This function can be handed app and a project and do
    whatever is required. This allows it to be run by
    another script if required
    """
    with pftextoutputs.PowerFactoryLogging(
            pf_app=app,
            add_handler=True,
            handler_level=logging.DEBUG,
            logger_to_use=logger,
            formatter=pftextoutputs.PFFormatter(
                '%(module)s: Line: %(lineno)d: %(message)s'  # format for printing stuff to the console
            )
    ) as pflogger:
        pf_main(app, project)


def pf_main(app=None, project=None):
    """
    Main function that runs if the script is executed using PowerFactory
    """
    project = app.GetActiveProject()
    if project is None:
        logger.error("No Active Project or passed project, Ending Script")
        return
    current_script = app.GetCurrentScript()
    execute_shc(app)
    # printmembers.print_power_factory_members(current_script,logger_function=logger.debug, print_desc=True)


def vs_main():
    '''
    Main function that runs if the script is executed outside of PowerFactory
    '''
    keys = ['Name', 'Longitude', 'Latitude']

    kml_dict = create_KML_file()

    '''
    folder1 = create_KML_folder(kml_dict['kmlDoc'], 'Folder1')
    kml_dict['document_element'].appendChild(folder1)

    folder2 = create_KML_folder(kml_dict['kmlDoc'], 'Folder2')
    kml_dict['document_element'].appendChild(folder2)
    '''
    csvdict = read_csv_to_dict('TestROAMES.csv', keys)
    next(csvdict)
    tour = generate_KML_tour(kml_dict, list(csvdict), 'My_Tour')
    
    kml_dict['document_element'].appendChild(tour)
    '''for line in csvdict:
        parentFolder1 = kml_dict['document_element'].getElementsByTagName('Folder').item(0)
        parentFolder2 = kml_dict['kml_element'].getElementsByTagName('Folder').item(1)  

        placemark_element = create_KML_placemarker(kml_dict, dict(line))
        print(dict(line))
        foo = line.get('Folder1')
        if foo == 'Yes':
            parentFolder2.appendChild(placemark_element)
        else:
            parentFolder1.appendChild(placemark_element)
    '''
    write_KML_file(kml_dict['kmlDoc'], 'ROAMES_tour.kml')
    print('finished')
    

def create_KML_file():
    '''
    Creates a basic KML file
    Returns the kml document (for writing) and the document element (for writing folders)
    '''
    kmlDoc = xml.dom.minidom.Document()  # basic kml document
  
    kml_element = kmlDoc.createElementNS('http://earth.google.com/kml/2.2', 'kml')
    kml_element.setAttribute('xmlns','http://earth.google.com/kml/2.2')
    kml_element.setAttribute('xmlns:gx', "http://www.google.com/kml/ext/2.2")
    kml_element.setAttribute('xmlns:kml', "http://www.opengis.net/kml/2.2")
    kml_element.setAttribute('xmlns:atom', "http://www.w3.org/2005/Atom")
    kml_element = kmlDoc.appendChild(kml_element)
    document_element = kmlDoc.createElement('Document')
    document_element = kml_element.appendChild(document_element)  # document element for creating folders

    return {'kmlDoc': kmlDoc, 'kml_element':kml_element, 'document_element':document_element}


def write_KML_file(kmlDoc, file_name):
    '''
    Writes the kml data to a given file name
    '''
    if file_name[-4:] != '.kml':  # add the file ending if it isnt already there
        file_name = file_name + '.kml'
    kmlFile = open(file_name, 'w')

    kmlFile.write(kmlDoc.toprettyxml(indent = '    '))  # write kmlDoc to the file


def create_KML_folder(kmlDoc, foldername):
    '''
    Creates a folder element to be added into the kmlDoc
    '''
    folderElement = kmlDoc.createElement('Folder')  # create the 'folder' tag
    folderElement.setAttribute('id', foldername)
    nameElement = kmlDoc.createElement('name')
    nameText = kmlDoc.createTextNode(foldername)
    nameElement.appendChild(nameText)
    folderElement.appendChild(nameElement)
    visibilityElement = kmlDoc.createElement('visibility')
    folderElement.appendChild(visibilityElement)
    visibilityValue = kmlDoc.createTextNode('1')  # set the visibility to off (0) as the default is on
    visibilityElement.appendChild(visibilityValue)

    return folderElement


def create_KML_placemarker(kmlDoc, data):
    '''
    Creates a placemark element given the data dictionary passed in
    data = {'Name': str , 'Latitude': float , 'Longitude: float, ...}
    '''
    placemark_element = kmlDoc['kmlDoc'].createElement('Placemark')  # tag containing all the information about a placemark

    # set the visibility to 1 (on)
    visibility_element = kmlDoc['kmlDoc'].createElement('visibility')
    placemark_element.appendChild(visibility_element)
    visibility_value = kmlDoc['kmlDoc'].createTextNode('1')
    visibility_element.appendChild(visibility_value)
    
    # create the Extended Data container and append to the Placemark instance 'placemarkElement'
    ext_element = kmlDoc['kmlDoc'].createElement('ExtendedData')
    placemark_element.appendChild(ext_element)  # appendChild is a way of oragnising the data (see the .kml file for this)

    name_element = kmlDoc['kmlDoc'].createElement('name')  # create the name element 'nameElement'
    name_text = kmlDoc['kmlDoc'].createTextNode(data['Name'])  # create a 'nameText' element and populate it with the name text from the row with the key 'Name'
    name_element.appendChild(name_text)  # append the 'nameText' element (containing the name text from the row) to 'nameElement'
    placemark_element.appendChild(name_element)  # append 'nameElement' to 'placemarkElement'

    for key in data:  # order is the list of column headers
        if key:
            data_element = kmlDoc['kmlDoc'].createElement('Data')
            data_element.setAttribute('name', key)
            value_element = kmlDoc['kmlDoc'].createElement('value')
            data_element.appendChild(value_element)
            value_text = kmlDoc['kmlDoc'].createTextNode(data[key])
            value_element.appendChild(value_text)
            ext_element.appendChild(data_element)

    point_element = kmlDoc['kmlDoc'].createElement('Point')  # coordinates must be stored in a 'point' element
    placemark_element.appendChild(point_element)
    coordinates = str(data['Longitude']) + ',' + str(data['Latitude'])  + ',0'
    coord_element = kmlDoc['kmlDoc'].createElement('coordinates')  # where the geographical coordinates are stored
    coord_element.appendChild(kmlDoc['kmlDoc'].createTextNode(coordinates))
    point_element.appendChild(coord_element)

    return placemark_element

def read_csv_to_list(file_name):
    if file_name[-4:] != '.csv':  # add the file ending if it isnt already there
        return False
    line_data = ()
    file_data = []
    read_file = open(file_name, 'r')
    for line in read_file:
        line_data = [element.strip() for element in line.split(',')]  # remove \ characters from line list
        file_data.append(line_data)
    if len(file_data) > 0:
        return file_data
    else:
        return False


def read_csv_to_dict(file_name, keys):
    if file_name[-4:] != '.csv':  # add the file ending if it isnt already there
        return False
    csvreader = csv.DictReader(open(file_name), keys)
    return csvreader


def execute_shc(app, terminal_element=None):
    shc = app.GetFromStudyCase('ComShc')

    # Basic Options
    logger.debug(dir(shc))
    shc.iopt_mde = 1  # short circuit calculation method
    shc.iopt_shc = '3psc'  # fault type
    shc.iopt_cur = 1  # calculate min or max short circuit currents
    shc.Rf = 0  # fault resistance
    shc.Xf = 0  # fault impedance
    if len(terminal_elements) == 1:
        shc.iopt_allbus = 0  # fault at user selection
        shc.shcobj = terminal_element
    else:
        shc.iopt_allbus = 2  # fault at all busbars
    shc.iopt_asc = True  # show output

    # Advanced Options
    if shc.iopt_mde == 1:  # advanced options for IEC60909 method
        shc.iopt_cdef = 0
        if shc.iopt_cdef == 1 or shc.iopt_cdef == 3:
            shc.cfac = 1
    elif shc.iopt_mde == 3:  # advanced options for complete method
        shc.ildfinnit = 1  # load flow initialisation
        if shc.ildfinnit == 0:
            shc.cfac_full = 1.1  # voltage factor c
            shc.ilgnLoad = 0  # loads
            shc.ilgnLneCap = 0  # capacitance of line
            shc.ilgnTrfMag = 0  # magnetising current of transformers
            shc.ilgnShnt = 0  # shunts/filters and SVS
    else:
        return "Unknown Method, Please Chose Either IEC60909 or Complete Method"
    shc.Execute()
    return "Short Circuit Successfully Executed"


def generate_KML_tour(kmlDoc, coords, tour_name='Tour'):
    tour_element = kmlDoc['kmlDoc'].createElement('gx:Tour')

    name_element = kmlDoc['kmlDoc'].createElement('name')
    name_value = kmlDoc['kmlDoc'].createTextNode(tour_name)
    name_element.appendChild(name_value)
    tour_element.appendChild(name_element)

    playlist_element = kmlDoc['kmlDoc'].createElement('gx:Playlist')
    

    if str(type(coords[0])) == "<class 'collections.OrderedDict'>":
        print('creating tour from list of dictionaries')

        # gx:duration
        # gx:flyToMode
        for element in coords:
            flyto_element = kmlDoc['kmlDoc'].createElement('gx:FlyTo')
            lookat_element = kmlDoc['kmlDoc'].createElement('LookAt')
            # gx:duration
            duration_element = kmlDoc['kmlDoc'].createElement('gx:duration')
            duration_value = kmlDoc['kmlDoc'].createTextNode('0.5')
            duration_element.appendChild(duration_value)
            flyto_element.appendChild(duration_element)

            # gx:flyToMode
            fly_to_mode_element = kmlDoc['kmlDoc'].createElement('gx:flyToMode')
            fly_to_mode_value = kmlDoc['kmlDoc'].createTextNode('smooth')
            fly_to_mode_element.appendChild(fly_to_mode_value)
            flyto_element.appendChild(fly_to_mode_element)
            
            # gx:horizFov
            horizon_FOV_element = kmlDoc['kmlDoc'].createElement('gx:horizFov')
            horizon_FOV_value = kmlDoc['kmlDoc'].createTextNode('60')
            horizon_FOV_element.appendChild(horizon_FOV_value)
            lookat_element.appendChild(horizon_FOV_element)

            # longitude
            longitude_element = kmlDoc['kmlDoc'].createElement('longitude')
            longitude_value = kmlDoc['kmlDoc'].createTextNode(element['Longitude'])
            longitude_element.appendChild(longitude_value)
            lookat_element.appendChild(longitude_element)

            # latitude
            latitude_element = kmlDoc['kmlDoc'].createElement('latitude')
            latitude_value = kmlDoc['kmlDoc'].createTextNode(element['Latitude'])
            latitude_element.appendChild(latitude_value)
            lookat_element.appendChild(latitude_element)

            # altitude
            altitude_element = kmlDoc['kmlDoc'].createElement('altitude')
            altitude_value = kmlDoc['kmlDoc'].createTextNode('0')
            altitude_element.appendChild(altitude_value)
            lookat_element.appendChild(altitude_element)

            # heading
            heading_element = kmlDoc['kmlDoc'].createElement('heading')
            heading_value = kmlDoc['kmlDoc'].createTextNode('-3')
            heading_element.appendChild(heading_value)
            lookat_element.appendChild(heading_element)

            # tilt
            tilt_element = kmlDoc['kmlDoc'].createElement('tilt')
            tilt_value = kmlDoc['kmlDoc'].createTextNode('15')
            tilt_element.appendChild(tilt_value)
            lookat_element.appendChild(tilt_element)

            # range
            range_element = kmlDoc['kmlDoc'].createElement('range')
            range_value = kmlDoc['kmlDoc'].createTextNode('687')
            range_element.appendChild(range_value)
            lookat_element.appendChild(range_element)
            
            # altitude mode
            altitude_mode_element = kmlDoc['kmlDoc'].createElement('gx:altitudeMode')
            altitude_mode_value = kmlDoc['kmlDoc'].createTextNode('relativeToSeaFloor')
            altitude_mode_element.appendChild(altitude_mode_value)
            lookat_element.appendChild(altitude_mode_element)

            # Wait
            wait_element = kmlDoc['kmlDoc'].createElement('gx:Wait')

            wait_duration_element = kmlDoc['kmlDoc'].createElement('gx:duration')
            wait_duration_value = kmlDoc['kmlDoc'].createTextNode('1.00')
            wait_duration_element.appendChild(wait_duration_value)

            wait_element.appendChild(wait_duration_element)


            flyto_element.appendChild(lookat_element)
            playlist_element.appendChild(flyto_element)
            playlist_element.appendChild(wait_element)

        tour_element.appendChild(playlist_element)
    elif str(type(coords[0])) == "<class 'list'>":
        print('creating tour form list of lists')
    else:
        print('invalid coordinate object')
        print(type(coords[0]))


    return tour_element




# def execute_ldf(app)
'''
TODO:
- create tour (from list of coordinates or csvreader dict)
- execute ldf (include as many variables)
- bokeh stuf


'''
if __name__ == '__main__':
    if PF:
        run_main()
    else:  # skips the PowerFactory setup if the script isnt being run through pf
        vs_main()
