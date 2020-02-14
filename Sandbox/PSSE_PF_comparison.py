import sys
# Python PowerFactory API
import powerfactory
import numpy as np
from operator import itemgetter
import pandas as pd
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
    Ensures a valid project is entered
    '''
    app.ClearOutputWindow()
    project = app.GetActiveProject()
    if project is None:
        activate_project(app)
    if project is None:
        logger.debug('No project active or passed through, ending script')
        return
    #  read the buses one by one and perform a complete fault, record Ikss and Skss for all busses and plot next to the database
    buses = app.GetCalcRelevantObjects('*.ElmTerm')
    
    execute_shc(app, buses)
    sh_data = read_shc_data(app, buses)
    pl_data, pl_N_minus_1_data = import_powerlink_spreadsheet(app)
    qld_sh_data = [bus for bus in sh_data if bus[0][0] == '4']  # get only the queensland buses
    qld_pl_N_minus_1_data = [bus for bus in pl_N_minus_1_data if bus[0][0] == '4']
    qld_pl_data = [bus for bus in pl_data if bus[0][0] == '4']
    
    sorted(qld_pl_data,key=itemgetter(0))  # sort in order of bus name
    sorted(pl_N_minus_1_data,key=itemgetter(0))
    sorted(qld_sh_data,key=itemgetter(0))

    logger.debug('powerlink N case: {}'.format(qld_pl_data))
    logger.debug('powerlink N-1 case: {}'.format(pl_N_minus_1_data))
    logger.debug('model: {}'.format(qld_sh_data))

    qld_sh_data = remove_non_duplicate_data(app, qld_pl_data, qld_sh_data) # powerlink data missing a lot of the busses that are in the model 

    logger.debug('new model: {}'.format(qld_sh_data))
    
    # compare_data(app, qld_pl_data, qld_sh_data)
    # logger.debug('Data comparison successful for N case')

    # now do N-1 case, dc (parameter: outserv) the element in critical contingency then run a fault on the target bus (parameter: shcobj)
    # slice names (lne = Br, trf = 2W or 3W, TE elements dont exist ignore them)

    
def execute_shc(app, buses):
    '''
    Perform a 3-phase short circuit on all buses in the model (complete method). Ignore all positive sequence data. 
    '''
    shc = app.GetFromStudyCase('ComShc')
    logger.debug('selected object is: {}'.format(app.GetCurrentSelection()))
    shc.iopt_mde = 3  # Complete fault mode
    shc.iopt_allbus = 2  # All busbars
    shc.Execute()


def read_shc_data(app, buses):
    '''
    Read the data gathered from short circuit analysis. Data from shorting all buses gave the same results as shorting each bus individually.
    '''
    results = []
    for i, bus in enumerate(buses):
        bus_ID = str(bus.loc_name)
        for ltr in bus_ID:  # filter out the area code of the bus and just get its unique ID
                if ltr == ' ':
                    space = bus_ID.index(ltr) + 1
                    bus_name = str(bus_ID[:space - 1])
        Ikss = float(bus.GetAttribute('m:Ikss'))
        Skss = float(bus.GetAttribute('m:Skss'))
        ip = float(bus.GetAttribute('m:ip'))
        real = float(bus.GetAttribute('m:R'))
        imaginary = float(bus.GetAttribute('m:X'))
        results.append([bus_name, Ikss, Skss, real, imaginary])
    return results


def import_powerlink_spreadsheet(app):
    '''
    Reads and formats the data from 'MinFL Summary for EQ-AEMO214 update.xlsx'. Spreadsheet had to be converted to a csv to be read by python due to UTF-8 encoding resitrictions.
    '''
    with open('Powerlink.csv', 'r') as pl_file:
        data = csv.reader(pl_file)
        next(data)
        next(data)

        pl_data_N = []
        pl_data_N_minus_1 = []  # worst variable name of the decade

        for record in data:
            bus_name = record[0]
            if bus_name == '':  # stop reading data if the next cell is empty
                break
            Ikss = float(record[12])
            Skss = float(record[13])
            Zth = record[17]
            
            real = float(Zth[1:5])
            for ltr in Zth:  # slice out the real and imaginary parts of the impedence
                if ltr == '+':
                    plus = Zth.index(ltr) + 1
                    imaginary = float(Zth[plus: plus + 5])
            pl_data_N.append([bus_name, Ikss, Skss, real, imaginary])

            critical_contingency = record[16]  # the element to be put out of service
            Ikss = float(record[14])
            Skss = float(record[15])
            Zth = record[18]
            
            real = float(Zth[1:5])
            for ltr in Zth:  # slice out the real and imaginary parts of the impedence
                if ltr == '+':
                    plus = Zth.index(ltr) + 1
                    imaginary = float(Zth[plus: plus + 5])
            # N-1 case: Ikss = 14, Skss = 15, Zth = 18
            # put the critical contingency out of service then run fault again

            pl_data_N_minus_1.append([bus_name, Ikss, Skss, real, imaginary, critical_contingency])

    return pl_data_N, pl_data_N_minus_1


def remove_non_duplicate_data(app, pl, sh):
    '''
    Removes the buses that arnt present in MinFL Summary for EQ-AEMO214 update.xlsx
    '''
    new_sh = []
    for model_bus in sh:
        for database_bus in pl:
            if model_bus[0] == database_bus[0]:
                new_sh.append(model_bus)
    return new_sh


def compare_data(app, pl, sh):
    '''
    calculate and display the powerfactory data compared to the PSSE data. the 'error' part can be removed if its easier to do in excel. 
    '''
    results = []
    results_file = open('Results.csv', 'w')
    results_file.write(' ,Powerfactory, , , , ,PSSE, , , , ,Error(%), , , , , ,\n')
    results_file.write('Name,Ikss(kA),Skss(MVA),real(), imag, ,Ikss,Skss,real, imag, ,Ikss,Skss,real, imag,\n')
    for i in range(len(pl)):
        error = [pl[i][0], (abs((sh[i][1]/pl[i][1]) - 1)*100), abs(((sh[i][2]/pl[i][2]) - 1)*100), abs(((sh[i][3]/pl[i][3]) - 1)*100), abs(((sh[i][4]/pl[i][4]) - 1)*100)]
        results_file.write('{},{},{},{},{}, ,{},{},{},{}, ,{},{},{},{},\n'.format(sh[i][0], sh[i][1], sh[i][2], sh[i][3], sh[i][4], pl[i][1], 
                                                                                pl[i][2], pl[i][3], pl[i][4], error[1], error[2], error[3], error[4]))



if __name__ == '__main__':
    run_main()