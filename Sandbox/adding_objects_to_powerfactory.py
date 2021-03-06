import sys
# Python PowerFactory API
import powerfactory as pf
import numpy as np
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
        main(app, project)


def main(app, project=None):
    """
    In Here is where you actually start doing stuff to the model.
    Iterate through stuff, call other functions. Etc.
    """
    project = app.GetActiveProject()
    if project is None:
        logger.error("No Active Project or passed project, Ending Script")
        return

    buses = app.GetCalcRelevantObjects('*.ElmTerm')

    graphics = app.GetActiveProject().GetContents('*.ElmNet', True)[0]  # first grid in the directory hierachy
    new_obj = graphics.CreateObject('ElmLne','Name')

    
    new_obj.loc_name = 'this_worked'
    new_obj.SetAttribute('bus1', buses[0])
    new_obj.SetAttribute('bus2', buses[1])
    app.PrintPlain('{} --- {}'.format(buses[0], buses[1]))
    # new_obj.bus1 = buses[0]
    # new_obj.bus2 = buses[1]

if __name__ == '__main__':
    run_main()
