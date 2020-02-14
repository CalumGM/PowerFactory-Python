import sys
# Python PowerFactory API
import powerfactory as pf

# Ergon Energy Helper Functions for Logging
sys.path.append(r'\\Ecasd01\WksMgmt\PowerFactory\Scripts\pfTextOutputs')
import pftextoutputs
# Ergon Energy Helper Functions for lots of stuff, inc Ecorp ID extraction
sys.path.append(r'\\Ecasd01\WksMgmt\PowerFactory\ScriptsDEV\pfSharedFunctions')
import pfsharedfunctions as pfsf
 
# Logging, A more efficient way to print
import logging
logger = logging.getLogger(__name__)

 
# Ergon Energy Helper Functions for getting members
sys.path.append(
    r"\\Ecasd01\WksMgmt\PowerFactory\ScriptsDEV\ShortPFScripts\PrintScripts\PrintPFObjectMembers"
)
import printPFObjectV2_descs as printmembers
 
# csv_reading library
import csv

# unsure why these are included
#from tkinter import filedialog

# from OS import *
import os


# from tinker import tk
#from tkinter import Tk

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
                '%(module)s: Line: %(lineno)d: %(message)s'
            )
    ) as pflogger:
        fault_level_summary_pia(app, project)
 
def fault_level_summary_pia(app, project=None):
    """
    Description
    """
    
    study_case_folder = app.GetProjectFolder("study")
    current_script = app.GetCurrentScript()

#do input checks
    #ask user to select study cases they are interested in running
    study_cases = multiple_modal_select_browser(app, study_case_folder.GetContents(), 'Please select the study cases you wish to run the PIA fault summary on.')
    if not study_cases:
        logger.error('No valid study cases selected. Ending Script')
        return
    for study_case in study_cases:
        if study_case.GetClassName() != 'IntCase':
            logger.error('One of the study cases selected is not valid. Ending script.')
            return
    study_case_folder = study_cases[0].fold_id

    #ask user to select which study case to copy the short_circuit_set from
    short_set_study_case = single_modal_select_browser(app, study_cases, 'Please select the study case which has the short circuit set you wish to use.')
    if not short_set_study_case:
        logger.error('Study Case not selected correctly. Ending Script.')
        return
    elif short_set_study_case.GetClassName() != 'IntCase':
        logger.error('Study Case not selected correctly. Ending Script')
        return
    base_short_set = short_set_study_case.GetContents('Short-Circuit Set.SetSelect')[0]
    logger.debug(base_short_set)

    #ask user which bus is associated with the 'new generator'
    new_gen_bus = single_modal_select_browser(app, base_short_set.All(), 'Please select which bus is associated with the NEW Generator')
    if not new_gen_bus:
        logger.error('Bus not selected correctly. Ending Script.')
        return
    elif new_gen_bus.GetClassName() != 'ElmTerm':
        logger.error('Bus not selected correctly. Ending Script')
        return
    logger.debug('new gen bus name is as follows:')
    logger.debug(new_gen_bus)
    

    #check for the New_Gen_ON variation
    #run gui to select the variations.
    variations = app.GetProjectFolder('scheme')
    new_gen_on = single_modal_select_browser(app, variations.GetContents(), 'Please select the variation which represents turning the NEW generator ON.')
    #logger.debug(new_gen_on.GetClassName())
    if not new_gen_on:
        logger.error('New_Gen_On variation not selected correctly. Ending Script.')
        return
    elif new_gen_on.GetClassName() != 'IntScheme':
        logger.error('New_Gen_On variation not selected correctly. Ending Script')
        return
    exist_async_on = single_modal_select_browser(app, variations.GetContents(), 'Please select the variation which represents turning the EXISTING asynchronous generators ON.')
    #logger.debug(exist_async_on.GetClassName())
    
    if not exist_async_on:
        logger.error('Existing_Gen_On variation not selected correctly. Ending Script')
        return        
    elif exist_async_on.GetClassName() != 'IntScheme':
        logger.error('Existing_Gen_On variation not selected correctly. Ending Script')
        return
    #logger.debug(operational_scenarios)
    #test_folder = variations.GetContents(r'Test*')
    #
    
    #check for the Exist_ASync_ON operational scenario
    
    #check if there is at least 1 bus of interest in the short-circuit-set

    #setup file location for .csv files
    output_file_folder = current_script.export_folder
   
    output_file_path = output_file_folder + r'\PIA_Results.csv'
    logger.debug(output_file_path)
    #write blank CSV to check if permissions exist and directory is correct
    fault_headers = ['']
    fault_summary = []
    save_csv_file(output_file_path, fault_headers, fault_summary)
#    root = Tk()
#    output_file_path = filedialog.askdirectory(title='Please create or select PIA directory') + r'/PIA_Results.csv'
#    root.destroy()
#    
    
    
    fault_dict = {}
#begin for loop to go through study cases
    for study_case in study_cases:
        base_study_case = study_case
        base_study_case.Activate()
#set short circuit set
        short_set = app.GetFromStudyCase('Short-Circuit Set.SetSelect')
        if short_set != base_short_set:
            short_set.Clear()
            short_set.AddRef(base_short_set.All()) #these couple lines basicaly ensures the set select used for each fault calculation is same.

#setup fault calculation settings
        
        #logger.debug(short_set)
        f_calc = app.GetFromStudyCase('ComShc')
        f_calc.iopt_mde = 1         #IEC 60909
        f_calc.iec_pub = 0         #2016
        f_calc.iopt_shc = '3psc'     #fault type
        f_calc.iopt_allbus = 0      #location user selection
        f_calc.shcobj = short_set  #perform fault calcs on every element in short circuit set

        #f_calc.Execute()

    #create new folder
        #activate appropriate operational scenarios per study case
        #create iether 3 or 4 study cases depending on if the new generator is synchonous
        pia_study_fold = study_case_folder.CreateObject('intFolder', base_study_case.loc_name + '_PIA_Summary')

        #run code for sync generators only
        study_sync_only = pia_study_fold.AddCopy(base_study_case, '1_Sync_Gens_Only')
        study_sync_only.Activate()
        exist_async_on.Deactivate()
        new_gen_on.Deactivate()
        f_calc.Execute()
        
        archive_fault_study_results(fault_dict, short_set, construct_tuple_key(base_study_case, study_sync_only.loc_name))
        #fault_summary.append(get_fault_study_results(short_set, base_study_case.loc_name, study_sync_only.loc_name, new_gen_bus))

        if current_script.new_sync == 1: #create extra study case if dealing with a sync gen
            study_sync_with_new = pia_study_fold.AddCopy(base_study_case, '2_Sync_Gens_With_New')
            study_sync_with_new.Activate()
            exist_async_on.Deactivate()
            new_gen_on.Activate()
            f_calc.Execute()
            
            archive_fault_study_results(fault_dict, short_set, construct_tuple_key(base_study_case, study_sync_with_new.loc_name))
            #fault_summary.append(get_fault_study_results(short_set, base_study_case.loc_name, study_sync_with_new.loc_name, new_gen_bus))

        #run code for all generators and new gen
        study_all_gens = pia_study_fold.AddCopy(base_study_case, '3_All_Gens_With_New')
        study_all_gens.Activate()
        exist_async_on.Activate()
        new_gen_on.Activate()
        f_calc.Execute()
        
        archive_fault_study_results(fault_dict, short_set, construct_tuple_key(base_study_case, study_all_gens.loc_name))
        #fault_summary.append(get_fault_study_results(short_set, base_study_case.loc_name, study_all_gens.loc_name, new_gen_bus))

        #run code for all generators minus the new
        study_all_gens_without_new = pia_study_fold.AddCopy(base_study_case, '4_All_Gens_Without_New')
        study_all_gens_without_new.Activate()
        exist_async_on.Activate()
        new_gen_on.Deactivate()
        f_calc.Execute()
        
        archive_fault_study_results(fault_dict, short_set, construct_tuple_key(base_study_case, study_all_gens_without_new.loc_name))
        #fault_summary.append(get_fault_study_results(short_set, base_study_case.loc_name, study_all_gens_without_new.loc_name, new_gen_bus))
        
        logger.debug(fault_dict)
        



#finish for loop
    
    #write output to csv
    for study_case in study_cases:
        fault_summary.extend(construct_PIA_tables(fault_dict, study_case, new_gen_bus, current_script.new_sync))

    logger.debug(output_file_path)
    save_csv_file(output_file_path, fault_headers, fault_summary)


#run fault study per study case and store results

#export results to CSV
def construct_tuple_key(network_scen, generator_scen):
    temp_tuple = (network_scen, generator_scen)
    return temp_tuple

def archive_fault_study_results(archive_dict, set_select, scenario_key):
    """ This function stores results from multiple fault studies (using a set select object), 
    and stores them appropriate to the tables rquired for the PIA. @connections_assessment
    """
    archive_dict[scenario_key] = {}
    for bus in set_select.All():
        archive_dict[scenario_key][bus] = bus.GetAttribute('m:Skss')

def construct_PIA_tables(archive_dict, study_case, new_gen_bus, new_sync=0):
    #construction necessary headers
    final_csv_array = []
    temp_row = []
    final_csv_array.append(['----', study_case.loc_name, '----'])
    final_csv_array.append(['*****', '*****', '*****'])
    temp_row.extend(['----', 'Synchronous Generation Fault Level'])
    temp_row.extend(['----', 'Total Fault Level'])
    temp_row.extend(['----', 'Fault Level Delta'])
    temp_row.extend(['----', 'Available Fault Level', '----'])
    final_csv_array.append(temp_row)
    temp_row = []
    temp_row = ['Site']
    temp_row.extend(['Fault Level (MVA)']*8)
    final_csv_array.append(temp_row)
    temp_row = []
    temp_row = ['']
    temp_row.extend(['Without '+ new_gen_bus.loc_name, 'With ' + new_gen_bus.loc_name]*4)
    final_csv_array.append(temp_row)
    temp_row = []

    #begin writing data
    temp_row.append(new_gen_bus.loc_name)
    temp_row.append(archive_dict[construct_tuple_key(study_case, '1_Sync_Gens_Only')][new_gen_bus])
    if new_sync == 1:
        temp_row.append(archive_dict[construct_tuple_key(study_case, '2_Sync_Gens_With_New')][new_gen_bus])
    else:
        temp_row.append(archive_dict[construct_tuple_key(study_case, '1_Sync_Gens_Only')][new_gen_bus])
    temp_row.append(archive_dict[construct_tuple_key(study_case, '4_All_Gens_Without_New')][new_gen_bus])
    temp_row.append(archive_dict[construct_tuple_key(study_case, '3_All_Gens_With_New')][new_gen_bus])
    #calc delta function
    temp_row.extend(calc_fault_delta(archive_dict, study_case, new_gen_bus))

    final_csv_array.append(temp_row)
    temp_row = []
    
    short_set = study_case.GetContents('Short-Circuit Set.SetSelect')[0]
    for bus in short_set.All():
        if bus != new_gen_bus:
            temp_row.append(bus.loc_name)
            temp_row.append(archive_dict[construct_tuple_key(study_case, '1_Sync_Gens_Only')][bus])
            if new_sync == 1:
                temp_row.append(archive_dict[construct_tuple_key(study_case, '2_Sync_Gens_With_New')][bus])
            else:
                temp_row.append(archive_dict[construct_tuple_key(study_case, '1_Sync_Gens_Only')][bus])
            temp_row.append(archive_dict[construct_tuple_key(study_case, '4_All_Gens_Without_New')][bus])
            temp_row.append(archive_dict[construct_tuple_key(study_case, '3_All_Gens_With_New')][bus])
            #calc delta function
            temp_row.extend(calc_fault_delta(archive_dict, study_case, bus))

            final_csv_array.append(temp_row)
            temp_row = []
    
    return final_csv_array

def calc_fault_delta(archive_dict, study_case, bus, new_sync = 0):
    return_list = []
    result1 = float(archive_dict[construct_tuple_key(study_case, '1_Sync_Gens_Only')][bus])
    result3 = float(archive_dict[construct_tuple_key(study_case, '3_All_Gens_With_New')][bus])
    result4 = float(archive_dict[construct_tuple_key(study_case, '4_All_Gens_Without_New')][bus])
    if new_sync ==1:
        result2 = float(archive_dict[construct_tuple_key(study_case, '2_Sync_Gens_With_New')][bus])
    else:
        result2 = result1
    return_list.append(result4 - result1)
    return_list.append(result3 - result2)
    return_list.append(2 * result1 - result4)
    return_list.append(2 * result2 - result3)

    return return_list
    




def append_table_headers(array, new_gen_bus, table_header):
    array.append(['----', table_header, '----'])
    array.append(['SITE', 'Fault Level (MVA)', 'Fault Level (MVA)'])
    array.append(['', 'Without '+ new_gen_bus.loc_name, 'With ' + new_gen_bus.loc_name])


def get_fault_study_results(set_select, network_scen, generator_scen, new_gen_bus):
    """ This function stores results from multiple fault studies (using a set select object), 
    and stores them appropriate to the tables rquired for the PIA. @connections_assessment
    """
    f_results = []
    f_results.append(network_scen)
    f_results.append(generator_scen)
    f_results.append(new_gen_bus.GetAttribute('m:Skss'))
    for bus in set_select.All():
        if bus != new_gen_bus:
            f_results.append(bus.GetAttribute('m:Skss'))
    return f_results

def save_csv_file(fPath, h, d):
    # Save the csv rows to file

    with open(fPath, 'w', newline='') as csvfile:
        cwriter = csv.writer(
            csvfile, delimiter=',', quotechar='"', quoting=csv.QUOTE_MINIMAL)
        heading_list = []
        for heading in h:
            heading_list.append(heading)
        cwriter.writerow(heading_list)
        for row in d:
            row_list = []
            for cell in row:
                row_list.append(cell)
            cwriter.writerow(row_list)

    return fPath

def single_modal_select_browser(app, input_list, title):
    """ This function allows the user to select a single object from a folder, 
    if a folder is selected it then allows the user to look in that folder and select
    """
    valid_selection = False

    while valid_selection == False:
        selection = app.ShowModalSelectBrowser(input_list, title)
        sel_len = len(selection)

        if sel_len == 0:
            return None
        elif sel_len >= 2:
            logger.error('Multiple scenarios selected. Please try again.')
            valid_selection = False
        elif selection[0].GetClassName() == 'IntFolder':
            input_list = selection[0].GetContents()
            valid_selection = False
        else:
            valid_selection = True


    return selection[0]

def multiple_modal_select_browser(app, input_list, title):
    """ This function allows the user to select a multiple objects from a folder, 
    if a folder is selected it then allows the user to look in that folder and select
    """
    valid_selection = False
    is_folder = False

    while valid_selection == False:
        selection = app.ShowModalSelectBrowser(input_list, title)
        sel_len = len(selection)

        if sel_len == 0:
            return None
        elif sel_len >= 2:
            for sel in selection:
                if sel.GetClassName() == 'IntFolder':
                    is_folder = True
            if is_folder:
                logger.error('Cannot select multiple objects if one is a folder. Please try again.')
                valid_selection = False
            else:
                valid_selection = True

        elif selection[0].GetClassName() == 'IntFolder':
            input_list = selection[0].GetContents()
            valid_selection = False
        else:
            valid_selection = True


    return selection




if __name__ == '__main__':
    run_main()
