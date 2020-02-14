import sys
# Python PowerFactory API
import powerfactory as pf
import numpy as np
import pandas as pd
from lxml import *
from pykml.factory_test import KML_ElementMaker as KML
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
        create_voltage_diff_plot(app, project)
 
def create_voltage_diff_plot(app, project=None):
    """
    In Here is where you actually start doing stuff to the model.
    Iterate through stuff, call other functions. Etc.
    """
    project = app.GetActiveProject()
    if project is None:
        logger.error("No Active Project or passed project, Ending Script")
        return

    main(app)
    current_script = app.GetCurrentScript()


def main(app):
	
    ldf = app.GetFromStudyCase('ComLdf')
    success = ldf.Execute()  # returns a 0 on success and >1 on fail
	
	
    if success == 0:
        app.PrintPlain('Name of all lines:')
        print_lines(app)
		
        app.PrintPlain('Length of L3-4:')
        print_line1(app)
		
        app.PrintPlain('Changing the length of L3-4:')
        modify_length(app)
        
        name_object = KML.name("Hello World!")



	
	
def print_lines(app):
	'''
	Print all the lines in the project
	'''
	lines = app.GetCalcRelevantObjects('*.ElmLne')
	for line in lines:
		app.PrintPlain(line.loc_name)
	app.PrintPlain('\n')
		
def modify_length(app):
	'''
	change the length of a given line
	'''
	modify_line = 'L3-4'
	new_length = 5
	AllObj = app.GetCalcRelevantObjects()
	try:
		lines = app.GetCalcRelevantObjects(modify_line + '.ElmLne')
		line = lines[0] # get the first line in a list of lines named 'L3-4'
		old_length = line.dline
		line.dline = new_length
		app.PrintPlain('Length of {} has been changed from {} to {}'.format(line.loc_name, old_length, line.dline))
	except IndexError:
		app.PrintPlain('couldnt find line')
	
	
def print_line1(app):
	'''
	print the first line
	'''
	first_line = 'L3-4'
	AllObj = app.GetCalcRelevantObjects()
	line = app.GetCalcRelevantObjects(first_line + '.ElmLne')
	try:
		app.PrintPlain('{}km'.format(line[0].dline))
	except IndexError:
		app.PrintPlain('couldnt find line\n')

		
if __name__ == '__main__':
    run_main()
