import sys
# Python PowerFactory API
import powerfactory
import numpy
import pandas
from bokeh.plotting import figure, show, output_file

# Ergon Energy Helper Functions for Logging
sys.path.append(r'\\Ecasd01\WksMgmt\PowerFactory\Scripts\pfTextOutputs')
import pftextoutputs
# Ergon Energy Helper Functions for lots of stuff, inc Ecorp ID extraction
sys.path.append(r'\\Ecasd01\WksMgmt\PowerFactory\ScriptsDEV\pfSharedFunctions')
import pfsharedfunctions

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
    app = powerfactory.GetApplication()
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
       main(app)

	
def main(app):
    '''
    Aim: Read the short circuit currents (Ikss) from the active study case and display relevant infomation
    '''
    app.ClearOutputWindow()
    project = app.GetActiveProject()
    if project is None:
    	activate_project(app)
    if project is None:
    	logger.debug('No project active or passed through, ending script')
        return
    buses = app.GetCalcRelevantObjects('*.ElmTerm')
    lines = app.GetCalcRelevantObjects('*.ElmLne')

    run_n_minus_one_fault(app, buses[0], lines[0])
    # load_flow(app)
    # short_circuit(app)
    # read_quasi_results(app)

	
def activate_project(app):
	user=app.GetCurrentUser()
	all_projects = user.GetContents('*',0)

	fundementals_project = [project for project in all_projects if project.loc_name == 'Fundamentals'][0]
	projects = fundementals_project.GetContents('*.IntPrj', 0)
	selection = app.ShowModalSelectBrowser(projects, 'Select A Project To Activate')
	selected_project = [project for project in projects if project.loc_name == selection[0].loc_name][0]

	success = app.ActivateProject('\cm233.IntUser\Fundamentals\PF-2-04-begin.IntPrj')
	logger.debug('failed {}'.format(success))
	logger.debug(app.GetActiveProject())  # why tf does this not work?

	
def load_flow(app):
	lines = app.GetCalcRelevantObjects('*.ElmLne')
	

	ldf = app.GetFromStudyCase('ComLdf')
	ldf.Execute()
		
	for line in lines:
		line_name = line.GetAttribute('loc_name')
		line_loading = line.GetAttribute('c:loading')
		line_loading = '{:.2f}'.format(line_loading)
		logger.debug('{:7}: {}%'.format(line_name, line_loading))

		
def short_circuit(app):
    buses = app.GetCalcRelevantObjects('*.ElmTerm')
    lines = app.GetCalcRelevantObjects('*.ElmLne')
    shc = app.GetFromStudyCase('ComShc')

    shc.iopt_mde = 3  # complete method
    shc.shcobj = buses[1]
    shc.iopt_allbus = 0

    lines[3].outserv = True
    logger.debug('{} is out of service'.format(lines[3]))
    shc.Execute()
    logger.debug('0 if no results: {}'.format(app.IsShcValid()))
	
    buses = app.GetCalcRelevantObjects('*.ElmTerm')
    short_bus = buses[1]

    Ikss = short_bus.GetAttribute('m:Ikss')
    logger.debug('{} : {:.2f}'.format(short_bus.loc_name, Ikss))

    connected_lines = [element for element in short_bus.GetConnectedElements(0,0,0) if element.GetClassName() == 'ElmLne']
    logger.debug('c: quantities exist: {}'.format(connected_lines))

def run_n_minus_one_fault(app, fault_bus, fault_part):
    fault_part.outserv = True  # take this part out of service
    logger.debug('{} is out of service'.format(fault_part))
    shc = app.GetFromStudyCase('ComShc')
    shc.iopt_mde = 3  # complete method
    shc.shcobj = fault_bus  # execute the short at this bus
    shc.iopt_allbus = 0
    shc.ildfinit = False  # set short circuit calculation assumptions
    shc.cfac_full = 1
    shc.ilngLoad = True
    shc.ilngLneCap = True
    shc.ilngTrfMag = True
    shc.ilngShnt = True
    shc.Execute()
    logger.debug('Ran fault on {}'.format(fault_bus))
    fault_part.outserv = False  # put the part back in service for the next fault

def read_quasi_results(app):
    study_case = app.GetActiveStudyCase()
    quasi_results = study_case.GetContents('Quasi-Dynamic Simulation AC', 0)[0]
    logger.debug(quasi_results)

    #loading the results file
    quasi_results.Load()

    number_of_rows = quasi_results.GetNumberOfRows()
    number_of_columns = quasi_results.GetNumberOfColumns()
    variables = []
    results = []
    for i in range(number_of_columns):
        for j in range(number_of_rows):
            x = '{:.3f}'.format(quasi_results.GetValue(j,i)[1])
            results.append(x)
        variables.append(results)
        results = []
    logger.debug(variables)
    '''
    external_data_directory = Get_Ext_Data_Dir(app.GetActiveProject())
    current_script = app.GetCurrentScript()
    output_folder = current_script.output_folder

    study_case_folder = app.GetProjectFolder("study")

    file_name = ext_data_dir + r'/' + output_folder + r'/' + study_case.loc_name + "_plots" + ".html"
    save(gridplot(variables, ncols=2, plot_width=600, plot_height=600, toolbar_location='right'))
    '''


if __name__ == '__main__':
	run_main()
