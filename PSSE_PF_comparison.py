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
    
    sh_data = n_case(app, buses)
    

    pl_data, pl_n_minus_one_data = import_powerlink_spreadsheet(app)
    qld_sh_data = [bus for bus in sh_data if bus[0][0] == '4']  # get only the queensland buses
    qld_pl_data = [bus for bus in pl_data if bus[0][0] == '4']
    qld_pl_n_minus_one_data = [bus for bus in pl_n_minus_one_data if bus[0][0] == '4']
    
    sorted(qld_sh_data,key=itemgetter(0))
    sorted(qld_pl_data,key=itemgetter(0))  # sort in order of bus name numerically
    sorted(qld_pl_n_minus_one_data,key=itemgetter(0))
    
    qld_sh_n_minus_one_data = n_minus_one_case(app, buses, qld_pl_n_minus_one_data)

    logger.debug('powerlink N case: {}'.format(qld_pl_data))
    logger.debug('powerlink N-1 case: {}'.format(qld_pl_n_minus_one_data))

    qld_sh_data = remove_non_duplicate_data(app, qld_pl_data, qld_sh_data) # powerlink data missing a lot of the busses that are in the model 

    logger.debug('new model: {}'.format(qld_sh_data))
    
    compare_data(app, qld_pl_data, qld_sh_data, 'N Case Results.csv')
    logger.debug('Data comparison successful for N case')

    compare_data(app, qld_pl_n_minus_one_data, qld_sh_n_minus_one_data, 'N-1 Case Results.csv')
    logger.debug('Data comparison successful for N-1 case')



def n_case(app, buses):
    execute_shc(app, buses)
    sh_data = read_shc_data(app, buses)  # n-case
    return sh_data


def n_minus_one_case(app, buses, pl):
    formated_pl = []
    all_fault_data = []
    logger.debug('\n \n')
    for index, record in enumerate(pl):
        critical_contingency = record[5]
        critical_contingency = critical_contingency.replace("'", '')
        if 'Br' in critical_contingency:
            critical_contingency = critical_contingency.replace('Br', 'lne')  # change the Br to lne
            record[5] = critical_contingency
            formated_pl.append(record)
            fault_bus = [bus for bus in buses if int(bus.loc_name[:6]) == int(record[0])][0]  # the bus that is to be shorted
            lines = app.GetCalcRelevantObjects('*.ElmLne')
            fault_part = [line for line in lines if line.loc_name == critical_contingency][0]  # the part to be taken out of service
            logger.debug('fault on bus: {} --- OOS: {}'.format(fault_bus, fault_part))
            fault_data = run_n_minus_one_fault(app, fault_bus, fault_part)[0]
        elif '2W' in critical_contingency:
            critical_contingency =  critical_contingency.replace('2W', 'trf')  # for all 2 winding transformers
            record[5] = critical_contingency
            formated_pl.append(record)
            fault_bus = [bus for bus in buses if int(bus.loc_name[:6]) == int(record[0])][0]
            transformers = app.GetCalcRelevantObjects('*.ElmTr2')
            critical_contingency_flipped = 'trf' + '_' + critical_contingency[10:15]  + '_' + critical_contingency[4:9] + '_' + critical_contingency[16]  # for transformers the otherway around
            fault_part = [transformer for transformer in transformers if transformer.loc_name == critical_contingency or transformer.loc_name == critical_contingency_flipped][0]
            logger.debug('fault on bus: {} --- OOS: {}'.format(fault_bus, fault_part))
            fault_data = run_n_minus_one_fault(app, fault_bus, fault_part)[0]
        elif '3W' in critical_contingency:
            critical_contingency =  critical_contingency.replace('3W', 'tr3')  # for all 3 winding transformers
            record[5] = critical_contingency
            formated_pl.append(record)
            fault_bus = [bus for bus in buses if int(bus.loc_name[:6]) == int(record[0])][0]
            transformers = app.GetCalcRelevantObjects('*.ElmTr3')
            fault_part = [transformer for transformer in transformers if transformer.loc_name == critical_contingency][0]
            logger.debug('fault on bus: {} --- OOS: {}'.format(fault_bus, fault_part))
            fault_data = run_n_minus_one_fault(app, fault_bus, fault_part)[0]
        all_fault_data.append(fault_data)
    return all_fault_data



def run_n_minus_one_fault(app, fault_bus, fault_part):
    '''
    Runs a fault cause for the given bus and de-activated component. Then reads the data and returns it ready to be written to excel
    '''
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
    results = []
    bus_ID = str(fault_bus.loc_name)
    for ltr in bus_ID:  # filter out the area code of the bus and just get its unique ID
        if ltr == ' ':
            space = bus_ID.index(ltr) + 1
            bus_name = str(bus_ID[:space - 1])
    Ikss = float(fault_bus.GetAttribute('m:Ikss'))
    Skss = float(fault_bus.GetAttribute('m:Skss'))
    ip = float(fault_bus.GetAttribute('m:ip'))
    real = float(fault_bus.GetAttribute('m:R'))
    imaginary = float(fault_bus.GetAttribute('m:X'))
    results.append([bus_name, Ikss, Skss, real, imaginary])   
    
    fault_part.outserv = False  # put the part back in service for the next fault
    return results


def execute_shc(app, buses):
    '''
    Perform a 3-phase short circuit on all buses in the model (complete method). Ignore all positive sequence data. 
    '''
    shc = app.GetFromStudyCase('ComShc')
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


def compare_data(app, pl, sh, file_name):
    '''
    calculate and display the powerfactory data compared to the PSSE data. the 'error' part can be removed if its easier to do in excel. 
    '''
    results = []
    results_file = open(file_name, 'w')
    results_file.write(' ,Powerfactory, , , , ,PSSE, , , , ,Error(%), , , , , ,\n')
    results_file.write('Name,Ikss(kA),Skss(MVA),real, imag, ,Ikss,Skss,real, imag, ,Ikss,Skss,real, imag,\n')
    for i in range(len(pl)):
        error = [pl[i][0], (abs((sh[i][1]/pl[i][1]) - 1)*100), abs(((sh[i][2]/pl[i][2]) - 1)*100), abs(((sh[i][3]/pl[i][3]) - 1)*100), abs(((sh[i][4]/pl[i][4]) - 1)*100)]
        results_file.write('{},{},{},{},{}, ,{},{},{},{}, ,{},{},{},{},\n'.format(sh[i][0], sh[i][1], sh[i][2], sh[i][3], sh[i][4], pl[i][1], 
                                                                                pl[i][2], pl[i][3], pl[i][4], error[1], error[2], error[3], error[4]))
    results_file.close()



if __name__ == '__main__':
    run_main()