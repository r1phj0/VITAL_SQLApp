import os
import pandas as pd
from datetime import datetime
from tkinter import *
from Diagnostic_Classes import add_to_dict, get_post_info, generate_schedules_db_format, \
    generate_schedules_of, generate_schedules_nf, Stop_DB_or_NF, get_ROT_info
import warnings

warnings.simplefilter(action='ignore')
# gets path to lookups.xlsx that is in a users PVS tools and templates folder on box in order to read in
pvs_sites_path = 'PVS Sites.xlsx'
lookups_path = 'lookups.xlsx'
frequencies = pd.read_excel(lookups_path.replace('lookups.xlsx', 'Frequency Code From Zero Bases.xlsx'),
                            sheet_name='FREQs')
standby_list = ['standby time', 'spotter', 'lunch', 'unassigned time', 'assigned to other duties', 'shell', 'fuel',
                'wash up']
main_dict_columns = ['Schedule Name', 'Stop Number', 'Arrive Time', 'Depart Time', 'Stop Name', 'Duration',
                     'Postalization Excess', 'Postalization Reason', 'Total x Stops', 'Total x Duration',
                     'Annual Trips', 'Annual Excess', 'Frequency Code', 'Tour', 'Vehicle', 'Begin Effective Date',
                     'End Effective Date', 'Times Route', 'Times Trip', 'Source File']


def ask_multiple_choice_question(prompt, options):
    root = Tk()
    root.lift()
    root.title('VITAL Database Info')
    root.geometry('120x130+300+200')

    if prompt == 'Which Analysis?':
        root.title('Analysis')
    root.attributes('-topmost', True)
    root.after_idle(root.attributes, '-topmost', False)
    if prompt:
        Label(root, text=prompt).pack()
    v = IntVar()
    for i, option in enumerate(options):
        Radiobutton(root, text=option, variable=v, value=i).pack(anchor="w")
    Button(root, text="Submit", command=root.destroy).pack()
    root.mainloop()
    return options[v.get()]


# uses the list of dataframes to calculate the annual excess and from that the excess staffing
def calc_staffing_excess(stop_dfs, summary):
    paid_hours = ['', '', '', '', '', '']
    paid_hours[0] = (summary['Annual Trips'] * summary['Paid Hours']).sum()
    # stop_dfs = [pvs_time, pdc_time, pvs_to_pdc, standby]
    excess_indices = ['PVS', 'P&DC', 'PVS to P&DC', 'Standby Time', 'PTF', 'Total']
    extra_mins = ['', '', '', '', '', '']
    extra_people = ['', '', '', '', '', '']
    index = 0
    for df in stop_dfs:
        if not df.empty:
            extra_mins[index] = stop_dfs[index]['Annual Excess'].sum()
        else:
            extra_mins[index] = 0
        index += 1
    # Mike says its 2080, WH rates says 1789 for PS 07 and 08 hourly rate
    index = 0
    for mins in extra_mins[:4]:
        extra_people[index] = (mins / 60 / 1789)
        index += 1
    ptf_ppl = 0
    for person in extra_people[:4]:
        ptf_ppl = ptf_ppl + (person * .2)
    extra_people[4] = ptf_ppl
    extra_people[5] = sum(extra_people[:5])
    all_staffing = {'Type': excess_indices, 'Excess Minutes': extra_mins,
                    'Excess Staffing': extra_people}
    return pd.DataFrame(all_staffing)


# takes the schedule objects created from the HTMLs and puts them into dicts and then dataframes
# stop name is either lists (pvs names, pdc names, unassigned time possibilities) or strings (standby, spotter)
# main dict columns is defined at the beginning of the file and is used for add_to_dict for column names
def diagnostics_from_htmls(all_VITAL_schedules, stop_name, pvs_names, pdc_names, main_dict_columns, format):
    if stop_name == pvs_names:
        stop_info_dict = {k: [] for k in (main_dict_columns[0:8] + ['Total PVS Stops', 'Total PVS Duration'] +
                                          main_dict_columns[10:20])}
        sheet_scenario, sort_by_column = 'Postalization', 'Postalization Excess'
    elif stop_name == pdc_names:
        stop_info_dict = {k: [] for k in (main_dict_columns[0:8] + ['Total P&DC Stops', 'Total P&DC Duration'] +
                                          main_dict_columns[10:20])}
        sheet_scenario, sort_by_column = 'Postalization', 'Postalization Excess'
    elif stop_name == 'lunch':
        stop_info_dict = {k: [] for k in (main_dict_columns[0:6] + ['Schedule Start Time', 'Paid Time Check',
                                                                    'Lunch Check', 'Paid Time'] +
                                          main_dict_columns[10:20])}
        sheet_scenario, sort_by_column = 'Other stops', 'Duration'
    elif stop_name == 'problem':
        if format != 'new':
            stop_info_dict = {k: [] for k in (['Status'] + main_dict_columns[0:6] + ['Postal Excess/Sched Start', 'Postal/Lunch Reason']
                                              + main_dict_columns[10:20])}
        else:
            stop_info_dict = {k: [] for k in (main_dict_columns[0:6] + ['Postal Excess/Sched Start', 'Postal/Lunch Reason']
                        + main_dict_columns[10:20])}
        sheet_scenario, sort_by_column = 'problem', 'Duration'
    else:
        stop_info_dict = {k: [] for k in (main_dict_columns[0:6] + main_dict_columns[10:20])}
        sheet_scenario, sort_by_column = 'Other stops', 'Duration'

    for schedule in all_VITAL_schedules:
        if not schedule.spotter_schedule:
            if sheet_scenario == 'Postalization':
                    for stop in schedule.stops:
                        # prints if a stop is greater than or equal to 30 for PVS/PDC, less than 15 for PVS,
                        # or has an issue (doesn't stop at PVS or PDC)
                        # if stop in the current tab names AND
                        # [(is a PDC stop and is either greater than or equal to 30 or has an issue) OR
                        # is a PVS stop and is less than 15 or greater than or equal to 30 or has an issue)]
                        if stop.stop_name in stop_name and \
                                (((stop.stop_name in schedule.pdc_names) and ((stop.duration > 30) or
                                                                              (stop.postalization_reason_code != '')))
                                 or ((stop.stop_name in schedule.pvs_names) and (((15 < stop.duration)) or
                                     (stop.postalization_reason_code != '')))):
                            if stop_name == pvs_names:
                                # filling in blank of the stop_info variable in the Stop with total pvs stops/pvs duration
                                stop.stop_info[8] = schedule.number_of_pvs_stops
                                stop.stop_info[9] = schedule.total_pvs_duration
                            elif stop_name == pdc_names:
                                # filling in blank of the stop_info variable in the Stop with total pdc stops/pdc duration
                                stop.stop_info[8] = schedule.number_of_pdc_stops
                                stop.stop_info[9] = schedule.total_pdc_duration
                            if stop not in schedule.problem_stops:
                                schedule.problem_stops.append(stop)
                            add_to_dict(stop_info_dict, stop.stop_info)
            # printing problem stops
            elif sheet_scenario == 'problem':
                if len(schedule.problem_stops) > 0:
                    for stop in schedule.problem_stops:
                        # this uses the stop_info variable of a lunch, pvs or pdc stop to add to the dict
                        if type(stop) is Stop_DB_or_NF:
                            if stop in schedule.lunch_stops or stop.lunch_stop:
                                if format == 'from db':
                                    add_to_dict(stop_info_dict, ([stop.status] + stop.lunch_info[0:7] + [stop.lunch_info[8]] +
                                                             stop.lunch_info[11:]))
                                elif format == 'new':
                                    add_to_dict(stop_info_dict, (stop.lunch_info[0:7] + [stop.lunch_info[8]] + stop.lunch_info[12:]))
                                else:
                                    add_to_dict(stop_info_dict, (stop.lunch_info[0:8] + stop.lunch_info[10:12] + stop.lunch_info[12:]))
                            else:
                                if stop.stop_name in [pvs_names + pdc_names]:
                                    stop.stop_info[11] = round(float(schedule.annual_trips * stop.postalization_time), 2)
                                if format == 'from db':
                                    add_to_dict(stop_info_dict, ([stop.status] + stop.stop_info[0:8] + stop.stop_info[10:20]))
                                elif format == 'old':
                                    add_to_dict(stop_info_dict, ['HCR'] + stop.stop_info[0:8] + stop.stop_info[10:])
                                else:
                                    add_to_dict(stop_info_dict,
                                                (stop.stop_info[0:8] + stop.stop_info[10:]))
                        else:
                            if format == 'from db':
                                # if the stop is not a Stop object that means it is pvs_to_pdc stop which is just a list of info
                                add_to_dict(stop_info_dict, ([schedule.status] + stop[0:8] + stop[10:20]))
                            else:
                                add_to_dict(stop_info_dict, (stop[0:8] + stop[10:20]))

            elif stop_name == 'lunch':
                if schedule.has_lunch:
                    # this prints the before and after lunch stops if they are different from each other
                    if len(schedule.lunch_stops) > 0:
                        for stop in schedule.lunch_stops:
                            if format == 'from db':
                                add_to_dict(stop_info_dict, (stop.lunch_info[0:9] + stop.lunch_info[10:22]))
                            else:
                                add_to_dict(stop_info_dict, (stop.lunch_info[0:9] + stop.lunch_info[11:]))
                    else:
                        # if the only problem with lunch is the actual lunch stop (not before and after), print info
                        if schedule.lunch_stop.lunch_reason != 'Good lunch':
                            if format == 'from db':
                                add_to_dict(stop_info_dict, schedule.lunch_stop.lunch_info[0:9] +
                                            schedule.lunch_stop.lunch_info[10:22])
                            else:
                                add_to_dict(stop_info_dict, schedule.lunch_stop.lunch_info[0:9] +
                                            schedule.lunch_stop.lunch_info[11:22])
            else:
                # this deals with unassigned, standby and spotter tabs
                for stop in schedule.stops:
                    # prints if stop_name is in the list or string given as a parameter, the duration isn't zero and the
                    # stop is not a lunch stop
                    if stop.stop_name.lower() in stop_name and stop.duration != 0 and stop != schedule.lunch_stop:
                        add_to_dict(stop_info_dict, (stop.stop_info[0:6] + stop.stop_info[10:20]))

    dataframe = pd.DataFrame(stop_info_dict)
    # lunch and problem tabs are sorted by schedule number and stop numbers
    if stop_name == 'lunch' or stop_name == 'problem':
        dataframe.sort_values(by=['Schedule Name', 'Stop Number'], ascending=[True, True], inplace=True)
    else:
        # sorting the sheet by postalization excess or duration GREATEST to least
        dataframe.sort_values(by=[sort_by_column], ascending=False, inplace=True)

    return dataframe


# method that converts information into dictionaries for printing for only the stops that are either
# from PVS to P&DC or from P&DC to PVS. looking for extra time
def inbetween_pvs_pdc_df(all_VITAL_schedules):
    inbetween_info_dict = {k: [] for k in ['Schedule Name', 'Stop Numbers', 'From', 'Depart Time', 'Arrive Time',
                                           'Duration', 'Annual Trips', 'Annual Excess', 'Frequency Code', 'Tour',
                                           'Vehicle', 'Source File']}
    for schedule in all_VITAL_schedules:
        # does not print any spotter schedule information because it would be extremely long since most dont stop at
        # the pdc between spotter stops
        if not schedule.spotter_schedule:
            # pvs to pdc and pdc to pvs are kept as a list of a list, prints the list if it exists
            if schedule.pvs_to_pdc_stop:
                add_to_dict(inbetween_info_dict, schedule.pvs_to_pdc_stop[0])
            if schedule.pdc_to_pvs_stop:
                add_to_dict(inbetween_info_dict, schedule.pdc_to_pvs_stop[0])

    inbetween_df = pd.DataFrame(inbetween_info_dict)
    inbetween_df.sort_values(by=['Duration'], ascending=False, inplace=True)
    return inbetween_df


# returns a dataframe that is a summary of the schedules. holds the information listed in the dict below
def one_page_summary(all_vital_schedules, format):
    if format == 'from db':
        one_page_summary_dict = {k: [] for k in ['Schedule Number', 'Status', 'Frequency', 'Annual Miles', 'Annual Hours',
                                             'MVO/TTO', 'Paid Hours', 'Lunch Duration', 'Standby Time', 'Spotter Time',
                                             'Begin Effective Date', 'End Effective Date', 'Annual Trips']}
    else:
        one_page_summary_dict = {k: [] for k in
                                 ['Schedule Number', 'Frequency', 'Annual Miles', 'Annual Hours',
                                  'MVO/TTO', 'Paid Hours', 'Lunch Duration', 'Standby Time', 'Spotter Time',
                                  'Begin Effective Date', 'End Effective Date', 'Annual Trips']}
    for schedule in all_vital_schedules:
        if schedule.has_lunch:
            lunch_duration = schedule.lunch_stop.duration
        else:
            lunch_duration = 0
        if format == 'from db':
            sched_list = [schedule.schedule_number, schedule.status, schedule.frequency, schedule.annual_miles,
                          schedule.annual_hours, schedule.vehicle, schedule.paid_time, lunch_duration,
                          schedule.standby_time, schedule.spotter_time, schedule.effective_date,
                          schedule.end_effective_date, schedule.annual_trips]
        else:
            sched_list = [schedule.schedule_number, schedule.frequency, schedule.annual_miles,
                          schedule.annual_hours, schedule.vehicle, schedule.paid_time, lunch_duration,
                          schedule.standby_time, schedule.spotter_time, schedule.effective_date,
                          schedule.end_effective_date, schedule.annual_trips]
        add_to_dict(one_page_summary_dict, sched_list)
    return pd.DataFrame(one_page_summary_dict)


# takes in various forms of information (htmls, excels, and which type of format to use) and from that gets
# the pvs and pdc names and calls the appropriate way to generate schedules (based on format)
# then prints those schedules to excel
def print_diagnostic(files, opt_sched_file, format, called_from):
    # old and new HTMLs get the postal info from lookups the same way, db can only get postal info after schedules file
    # is found and read in
    # (format, opt_sched_file, postal_facility, html_files, called_from
    if format != 'from db':
        # new htmls called from optimizer aka files are not in HTML Check, given by app
        if called_from == 'from_optimizer' and format != 'old':
            pvs_names, pdc_names, pvs_and_pdc_names, postal_facility = get_post_info(format, 'opt_sched_file',
                                                                                     'post_facility', files[0],
                                                                                     called_from)
        # old htmls called from opti app so using Schedules file as finder for pvs/pdc names
        else:
            pvs_names, pdc_names, pvs_and_pdc_names, postal_facility = get_post_info(format, opt_sched_file,
                                                                                     'post_facility', 'html_files',
                                                                                     called_from)
        # new html format uses the Stop_NF and schedule_NF classes vs Stop_OF and schedule_OF for old html format
        # old and new html format classes can be found in Diagnostic_Classes file
        if format == 'new':
            all_vital_schedules = generate_schedules_nf(files, pvs_names, pdc_names, pvs_and_pdc_names)
        elif format == 'old':
            all_vital_schedules = generate_schedules_of(pd.read_excel(opt_sched_file,
                                                                      sheet_name='As Read Schedules',
                                                                      converters={'Arrive Time': str,
                                                                                  'Depart Time': str}),
                                                        pvs_names, pdc_names)
        postal_facilities = [postal_facility]
    # if format is from VITAL database, find all postal facilities in file (set this way so will print more than 1 site)
    else:
        # looking for schedules and stops files and reading them in as dataframes
        if called_from == 'from_optimizer':
            # looks for schedules file without ignoring 'zero base' because files are most likely from box if coming
            # from the optimizer so Zero Bases 202X will be in the path name
            schedules_file = pd.read_excel([file for file in files if 'schedules' in file.lower()][0])
            stops_file = pd.read_excel([file for file in files if 'stops' in file.lower()][0])
        else:
            schedules_file = pd.read_excel([file for file in files if 'schedules' in file.lower() and 'zero base' not in file.lower()][0])
            stops_file = pd.read_excel([file for file in files if 'stops' in file.lower() and 'zero base' not in file.lower()][0])
        postal_facilities = list(set(schedules_file['SITE_NAME']))
        postal_facilities.sort()
    for site in postal_facilities:
        rot_file = pd.DataFrame()
        if rot_file.empty:
            xl_files = []
            for file in os.listdir('HTML Check/'):
                if file.endswith('.xls') or file.endswith('.xlsx') or file.endswith('.XLSX'):
                    xl_files.append(str('HTML Check/' + file))
            try:
                rot_file = pd.read_excel([file for file in xl_files if 'rot' in file.lower()][0])
            except:
                file_path = g.fileopenbox("Choose ROT File")
                try:
                    rot_file = pd.read_excel(file_path)
                except:
                    print('You did not choose the ROT file')
                    quit()
        # all vital schedules already created if it is new or old HTML format
        # so this creates them if it's from the VITAL database
        if format == 'from db':
            # schedules file is needed for postal facility before getting pvs/pdc names
            pvs_names, pdc_names, pvs_and_pdc_names, postal_facility = get_post_info(format, 'opt_sched_file', site,
                                                                                     'html_files', called_from)
            site_specific_schedules = schedules_file[(schedules_file['SITE_NAME'] == site)]
            site_specific_stops = stops_file[(stops_file['SITE_NAME'] == site)]
            # generating vital schedules with the DB format classes that are in this file
            all_vital_schedules = generate_schedules_db_format(site_specific_schedules, site_specific_stops, pvs_names,
                                                               pdc_names)
        # minneapolis, st paul, etc use mailer facility as a pvs/pdc name, so if its in that list, it is
        # removed so that all the MAILER FACILITY stops do not end up in the Unassigned Time tab of the diagnostic
        unassigned_list = ['assigned to other duties', 'mailer facility', 'on demand']
        if 'MAILER FACILITY' in (pvs_names + pdc_names):
            unassigned_list.remove('mailer facility')
        # below is a line in case you want to separate TTO or MVO schedules
        # all_vital_schedules = [x for x in all_vital_schedules if x.vehicle == 'MVO']
        # must do all pvs/pdc stops first so that they can add all problem stops before problem tab
        pvs_to_pdc = inbetween_pvs_pdc_df(all_vital_schedules)
        pvs_stops = diagnostics_from_htmls(all_vital_schedules, pvs_names, pvs_names, pdc_names, main_dict_columns, format)
        pdc_stops = diagnostics_from_htmls(all_vital_schedules, pdc_names, pvs_names, pdc_names, main_dict_columns, format)
        problem = diagnostics_from_htmls(all_vital_schedules, 'problem', pvs_names, pdc_names, main_dict_columns, format)
        lunch = diagnostics_from_htmls(all_vital_schedules, 'lunch', pvs_names, pdc_names, main_dict_columns, format)
        standby_stops = diagnostics_from_htmls(all_vital_schedules, 'standby time', pvs_names, pdc_names,
                                               main_dict_columns, format)
        spotter = diagnostics_from_htmls(all_vital_schedules, 'spotter', pvs_names, pdc_names, main_dict_columns, format)
        unassigned_stops = diagnostics_from_htmls(all_vital_schedules, unassigned_list, pvs_names, pdc_names,
                                                  main_dict_columns, format)
        summary = one_page_summary(all_vital_schedules, format)
        # [pvs_time, pdc_time, pvs_to_pdc, standby], summary for total staffing
        excess_staffing = calc_staffing_excess([pvs_stops, pdc_stops, pvs_to_pdc, standby_stops], summary)
        try:
            zero_ROT = get_ROT_info(rot_file, pvs_names, pdc_names, site_specific_stops)
        except:
            zero_ROT = pd.DataFrame()
            print("Couldn't analyze ROT data")
        stop_dataframes_final = [problem, standby_stops, unassigned_stops, lunch, pvs_to_pdc, pvs_stops, pdc_stops, zero_ROT,
                                 excess_staffing, spotter, summary]
        sheet_names = ['All Problem Stops', 'Standby Time', 'Unassigned Time', 'Lunch Time', 'PVS to PDC Time',
                       'PVS Stops', 'P&DC Stops','ROT', 'Excess Staffing', 'Info - Spotter Time', 'Info - Summary']
        # removing unassigned time if there are no stops
        if unassigned_stops.empty:
            stop_dataframes_final = [problem, standby_stops, lunch, pvs_to_pdc, pvs_stops, pdc_stops, zero_ROT,
                                 excess_staffing, spotter, summary]
            sheet_names.remove('Unassigned Time')
        writer1 = pd.ExcelWriter((str("DataFiles/" + site) + ' Diagnostic ' + str(datetime.today().strftime('%m.%d.%y'))
                                  + '.xlsx'), engine='xlsxwriter')

        # variable for iterating
        a = 0

        # loop that goes through the list of dataframes per employee and prints each to a different sheet
        # the sheet name is gotten from the initial list of employee ids read in from final staffing "staffing" tab
        while a < len(stop_dataframes_final):
            # temp variable to hold each employees dataframe
            df_new = stop_dataframes_final[a]
            sheet_name = sheet_names[a]
            df_new.to_excel(writer1, sheet_name=sheet_name, index=False)
            # increase the variable to continue iterating
            a += 1

        # must save the writer object or else nothing will actually print
        writer1.save()


if __name__ == "__main__":
    # list of any .htm or .html files in the HTML Check folder
    # html_files = []
    # for file in os.listdir('HTML Check/'):
    #     if file.endswith('.htm') or file.endswith('.html'):
    #         html_files.append(str('HTML Check/' + file))
    # list of any .xls or .xlsx files in the HTML Check folder
    db_files = []
    for file in os.listdir('HTML Check/'):
        if file.endswith('.xls') or file.endswith('.xlsx') or file.endswith('.XLSX'):
            db_files.append(str('HTML Check/' + file))
    # analysis = ask_multiple_choice_question('Which Data?', ['Old', 'New', 'VITAL DB'])
    # analysis = 'VITAL DB'
    # db_files = ['HTML Check/']
    called_from = 'diagnostic.py'
    # if analysis == 'Old':
    #     # finds the schedules file for old HTML format to generate schedules
    #     opt_sched_file = [file for file in db_files if 'Schedules' in file]
    #     print_diagnostic(html_files, opt_sched_file[0], 'old', called_from)
    # elif analysis == 'New':
    #     print_diagnostic(html_files, '', 'new', called_from)
    # elif analysis == 'VITAL DB':
    print_diagnostic(db_files, '', 'from db', called_from)