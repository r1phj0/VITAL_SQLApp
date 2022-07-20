import os
from bs4 import BeautifulSoup
import pandas as pd
from tkinter import *
from datetime import datetime
import dateparser
import googlemaps
import math

gmaps = googlemaps.Client(key='AIzaSyDzJbrTuWXDVK4NiNzOIANVBWR3net2CsM')

pvs_sites_path = 'PVS Sites.xlsx'
lookups_path = 'lookups.xlsx'
lookups = pd.read_excel(lookups_path, sheet_name='PVS Lookup')
frequencies = pd.read_excel(lookups_path.replace('lookups.xlsx', 'Frequency Code From Zero Bases.xlsx'),
                            sheet_name='FREQs')
standby_list = ['standby time', 'spotter', 'lunch', 'unassigned time', 'assigned to other duties', 'shell', 'fuel']


# simple method that takes in a dictionary and list and adds everything from the list to the appropriate place
# in the dictionary (this is so that it is easily convertible to pandas for printing)
def add_to_dict(new_dict, additions):
    for x in new_dict:
        new_dict[x].append(additions[list(new_dict.keys()).index(x)])


def get_ROT_info(site_rot, site_pvs_names, site_pdc_names, stops):

    site_rot[['ORIG_ADDR1', 'ORIG_ADDR2', 'ORIG_CITY', 'ORIG_STATE']] = site_rot[
        ['ORIG_ADDR1', 'ORIG_ADDR2', 'ORIG_CITY', 'ORIG_STATE']].astype(str)
    site_rot['MILEAGE_NBR'].fillna(0, inplace=True)

    # get zero mileage records
    zero_mileage = site_rot[site_rot['MILEAGE_NBR'] == 0]
    zero_mileage.reset_index(drop=True, inplace=True)
    final_zero_mileage_df = pd.DataFrame(
        columns=['ORIG_NASS', 'ORIG_NAME', 'ORIG_ADDRESS', 'DEST_NASS', 'DEST_NAME', 'DEST_ADDRESS', 'MILEAGE_NBR',
                 'DRIVE_TIME', 'SCHEDULES'])
    # add records to final dataframe with all necessary info
    for i in range(zero_mileage.shape[0]):
        origin = zero_mileage.loc[i, 'ORIG_NAME']
        destination = zero_mileage.loc[i, 'DEST_NAME']
        if (origin in site_pvs_names and destination in site_pdc_names) or (
                origin in site_pdc_names and destination in site_pvs_names):
            continue
        else:
            place = final_zero_mileage_df.shape[0] + 1
            final_zero_mileage_df.loc[place, 'ORIG_NASS'] = zero_mileage.loc[i, 'ORIG_NASS']
            final_zero_mileage_df.loc[place, 'ORIG_NAME'] = origin
            if str(zero_mileage.loc[i, 'ORIG_ADDR2']) == 'nan':
                final_zero_mileage_df.loc[place, 'ORIG_ADDRESS'] = str(zero_mileage.loc[i, 'ORIG_ADDR1']) + ', ' + \
                                                                   str(zero_mileage.loc[
                                                                           i, 'ORIG_CITY']) + ', ' + str(
                    zero_mileage.loc[i, 'ORIG_STATE'])
            else:
                final_zero_mileage_df.loc[place, 'ORIG_ADDRESS'] = str(
                    zero_mileage.loc[i, 'ORIG_ADDR1']) + ' ' + str(zero_mileage.loc[i, 'ORIG_ADDR2']) + ', ' + str(
                    zero_mileage.loc[i, 'ORIG_CITY']) + ', ' + str(zero_mileage.loc[i, 'ORIG_STATE'])
            final_zero_mileage_df.loc[place, 'DEST_NASS'] = zero_mileage.loc[i, 'DEST_NASS']
            final_zero_mileage_df.loc[place, 'DEST_NAME'] = destination
            if str(zero_mileage.loc[i, 'DEST_ADDR2']) == 'nan':
                final_zero_mileage_df.loc[place, 'DEST_ADDRESS'] = str(zero_mileage.loc[i, 'DEST_ADDR1']) + ', ' + \
                                                                   str(zero_mileage.loc[
                                                                           i, 'DEST_CITY']) + ', ' + str(
                    zero_mileage.loc[i, 'DEST_STATE'])
            else:
                final_zero_mileage_df.loc[place, 'DEST_ADDRESS'] = str(
                    zero_mileage.loc[i, 'DEST_ADDR1']) + ' ' + str(zero_mileage.loc[i, 'DEST_ADDR2']) + ', ' + str(
                    zero_mileage.loc[i, 'DEST_CITY']) + ', ' + str(zero_mileage.loc[i, 'DEST_STATE'])
            final_zero_mileage_df.loc[place, 'MILEAGE_NBR'] = zero_mileage.loc[i, 'MILEAGE_NBR']
            final_zero_mileage_df.loc[place, 'SCHEDULES'] = 'Not used in schedules'

    # looking for missing routes using time_routes file
    xl_files = []
    for file in os.listdir('HTML Check/'):
        if file.endswith('.xls') or file.endswith('.xlsx') or file.endswith('.XLSX'):
            xl_files.append(str('HTML Check/' + file))

    time = pd.read_excel([file for file in xl_files if 'times' in file.lower()][0])

    # set up key between the two files
    site_rot['KEY'] = site_rot['ORIG_NASS'] + site_rot['DEST_NASS']
    time['KEY'] = time['ORIG_NASS'] + time['DEST_NASS']
    stops[['LINE1_ADDR', 'LINE2_ADDR', 'CITY_NAME', 'STATE_ID']] = stops[
        ['LINE1_ADDR', 'LINE2_ADDR', 'CITY_NAME', 'STATE_ID']].astype(str)

    # routes present in the ROT table file
    known_routes = set(list(site_rot['KEY']))

    # set up dataframe for missing routes
    missing_routes = pd.DataFrame(
        columns=['SCH_NBR', 'ORIG_NASS', 'ORIG_NAME', 'ORIG_ADDRESS', 'DEST_NASS', 'DEST_NAME', 'DEST_ADDRESS',
                 'MILEAGE_NBR', 'DRIVE_TIME', 'ROUTE', 'KEY', 'AFFECTED_SCHEDULES'])

    # check whether a route is accounted for - if not, adds to the missing df with info
    for i in range(time.shape[0]):
        route = time.loc[i, 'KEY']
        if route in known_routes:
            continue
        else:
            origin_nass = time.loc[i, 'ORIG_NASS']
            dest_nass = time.loc[i, 'DEST_NASS']
            sch_nbr = str(time.loc[i, 'SCH_SCHED_NBR'])
            sch_stops = stops[stops['SCH_SCHED_NBR'] == sch_nbr]
            name_dict = pd.Series(sch_stops['STOP_NAME'].values, index=sch_stops['NASS_CD']).to_dict()
            end = missing_routes.shape[0] + 1
            try:
                origin_name = name_dict[origin_nass]
            except:
                origin_name = "Not Found"
            try:
                dest_name = name_dict[dest_nass]
            except:
                dest_name = "Not Found"
            sch_stops['STOP_NAME'] = sch_stops['STOP_NAME'].map(str)
            sch_stops['STOP_NAME'] = sch_stops['STOP_NAME'].apply(lambda x: x.upper())
            new_sch_stops = sch_stops[sch_stops['STOP_NAME'] != 'SPOTTER']
            new_sch_stops = new_sch_stops[new_sch_stops['STOP_NAME'] != 'LUNCH']
            new_sch_stops['STOP_NAME'].fillna('PVS', inplace=True)
            new_sch_stops.reset_index(inplace=True)
            for i in range(new_sch_stops.shape[0]):
                if str(new_sch_stops.loc[i, 'LINE2_ADDR']) == 'nan':
                    new_sch_stops.loc[i, 'ADDRESS'] = str(
                        new_sch_stops.loc[i, 'LINE1_ADDR']) + ', ' + str(
                        new_sch_stops.loc[i, 'CITY_NAME']) + ', ' + str(
                        new_sch_stops.loc[i, 'STATE_ID'])
                else:
                    new_sch_stops.loc[i, 'ADDRESS'] = str(
                        new_sch_stops.loc[i, 'LINE1_ADDR']) + ' ' + str(
                        new_sch_stops.loc[i, 'LINE2_ADDR']) + ', ' + str(
                        new_sch_stops.loc[i, 'CITY_NAME']) + ', ' + str(new_sch_stops.loc[i, 'STATE_ID'])
            address_dict = pd.Series(new_sch_stops['ADDRESS'].values, index=new_sch_stops['NASS_CD']).to_dict()
            if not origin_name == "Not Found":
                origin_add = address_dict[origin_nass]
            else:
                origin_add = ""
                print(f'ORIGIN NOT FOUND for {sch_nbr} - {origin_nass}')
            if not dest_name == "Not Found":
                dest_add = address_dict[dest_nass]
            else:
                dest_add = ""
                print(f'DESTINATION NOT FOUND for {sch_nbr} - {dest_nass}')
            missing_routes.loc[end, 'SCH_NBR'] = sch_nbr
            missing_routes.loc[end, 'ORIG_NASS'] = origin_nass
            missing_routes.loc[end, 'ORIG_NAME'] = origin_name
            missing_routes.loc[end, 'ORIG_ADDRESS'] = origin_add
            missing_routes.loc[end, 'DEST_NASS'] = dest_nass
            missing_routes.loc[end, 'DEST_NAME'] = dest_name
            missing_routes.loc[end, 'DEST_ADDRESS'] = dest_add
            missing_routes.loc[end, 'MILEAGE_NBR'] = 'Missing'
            missing_routes.loc[end, 'ROUTE'] = route

    # gets unique list of missing routes
    unique_routes = set(list(missing_routes['ROUTE']))
    affected_routes = {}

    # loops through unique routes and gets all relevant schedule numbers associated with route
    for value in unique_routes:
        schedules = []
        temp_df = missing_routes[missing_routes['ROUTE'] == value]
        temp_df.reset_index(drop=True, inplace=True)
        for i in range(temp_df.shape[0]):
            sch = temp_df.loc[i, 'SCH_NBR']
            schedules.append(int(sch))
        schedules = set(schedules)
        affected_routes[value] = sorted(schedules)

    # drops duplicate routes
    missing_routes.drop_duplicates(subset=['ORIG_NAME', 'DEST_NAME'], inplace=True)
    missing_routes.reset_index(drop=True, inplace=True)

    # sort missing routes by points so that routes with same stops are next to each other
    # add a ordered route key used to sort
    for i in range(missing_routes.shape[0]):
        points = []
        o_nass = missing_routes.loc[i, 'ORIG_NASS']
        d_nass = missing_routes.loc[i, 'DEST_NASS']
        points.append(o_nass)
        points.append(d_nass)
        points.sort()
        key = str(points[0]) + str(points[1])
        missing_routes.loc[i, 'KEY'] = key

    # sort dataframe
    missing_routes.sort_values(by=['KEY'], inplace=True)

    # clean dfs and combine to get final output
    missing_routes['SCHEDULES'] = missing_routes['ROUTE'].map(affected_routes)
    missing_routes = missing_routes[
        ['ORIG_NASS', 'ORIG_NAME', 'ORIG_ADDRESS', 'DEST_NASS', 'DEST_NAME', 'DEST_ADDRESS', 'MILEAGE_NBR',
         'DRIVE_TIME', 'SCHEDULES']]
    final_df = pd.concat([final_zero_mileage_df, missing_routes])

    # add mileage and drive time
    final_df.reset_index(drop=True, inplace=True)
    for i in range(final_df.shape[0]):
        address1 = final_df.loc[i, 'ORIG_ADDRESS']
        address2 = final_df.loc[i, 'DEST_ADDRESS']
        try:
            dist = gmaps.distance_matrix(address1, address2)
            temp = dist["rows"][0]["elements"][0]
            assert temp["status"] == "OK"
            dist_output = temp["distance"]["value"]
            dur_output = temp["duration"]["value"]
            dur = max(1, int(math.ceil(dur_output / 60)))
            five_dur = int(math.ceil(dur/5.0))*5
            distance = max(0, round(dist_output / 1609, 1))
            final_df.loc[i, 'MILEAGE_NBR'] = distance
            final_df.loc[i, 'DRIVE_TIME'] = str(five_dur) + "m"
        except:
            continue

    final_df.reset_index(drop=True, inplace=True)

    return final_df


class All_schedules:
    def __init__(self, schedule_unique_ind, duplicates):
        self.schedule_unique_ind = schedule_unique_ind
        self.duplicates = duplicates


# takes the datetime objects of depart and arrive time
def calculate_duration(depart_time, arrive_time, stop_or_sched):
    if stop_or_sched == 'stop':
        if depart_time == arrive_time:
            duration = 0
        elif (depart_time - arrive_time).total_seconds() / 60 < 1:
            duration = ((depart_time - arrive_time).total_seconds()/60) + 1440
        else:
            duration = (depart_time - arrive_time).total_seconds()/60
    else:
        #
        if (arrive_time - depart_time).total_seconds() / 60 / 60 > 0:
            duration = (arrive_time - depart_time).total_seconds()/60/60
        else:
            depart_time = depart_time.replace(day=2)
            duration = (depart_time - arrive_time).total_seconds()/60/60
        if duration > 12:
            depart_time = depart_time.replace(day=1)
            arrive_time = arrive_time.replace(day=2)
            duration = (depart_time - arrive_time).total_seconds() / 60 / 60
            if duration < 0:
                duration = (arrive_time - depart_time).total_seconds() / 60 / 60
    return duration


# Class that creates a frame for each schedule that has a button with the schedule number and
# 3 radio buttons for delete, mvo and tto
class Checkbar(Frame):
    def __init__(self, parent=Frame, sched_row_df=[], anchor=W):
        Frame.__init__(self, parent)
        schedule_set = sched_row_df[0]
        self.schedule_unique_ind = sched_row_df[1]
        self.schedule_num = str(schedule_set['SCH_SCHED_NBR'])
        self.effective_date = schedule_set['SCH_EFFECT_DTM']
        self.effective_date_text = str(' | Effective Date: ' + schedule_set['SCH_EFFECT_DTM'])
        self.end_effective_date_text = str(' | End Effective Date: ' + schedule_set['END_DT'])
        self.end_effective_date = schedule_set['END_DT']
        self.start_time_text = str(' | Start Time: ' + schedule_set['START_TIME'])
        self.start_time = schedule_set['START_TIME']
        self.end_time_text = str(' | End Time: ' + schedule_set['END_TIME'])
        self.end_time = schedule_set['END_TIME']
        self.status_text = str(' | Status: ' + schedule_set['DECODE_DESC'])
        self.status = schedule_set['DECODE_DESC']
        self.tour_text = str(' | Tour: ' + str(schedule_set['TOUR_NBR']))
        self.tour = schedule_set['TOUR_NBR']
        self.mileage_text = str(' | Mileage: ' + str(schedule_set['MILEAGE_NBR']))
        self.mileage = schedule_set['MILEAGE_NBR']
        self.stop_count_text = str(' | Stops: ' + str(schedule_set['TOT_STOP_CNT']))
        self.stop_count = schedule_set['TOT_STOP_CNT']
        self.vars = []
        attributes = [self.start_time_text, self.effective_date_text, self.status_text, self.schedule_num]
        picks = ['Include', 'Exclude']
        counter = 1
        while counter < 3:
            var = IntVar()
            var.set(counter-1)
            chk = Radiobutton(self, text=picks[counter-1], var=var, value=1, command=self.change)
            chk.pack(side=LEFT, anchor=anchor, expand=YES)
            self.vars.append(var)
            counter += 1
        for attribute in attributes:
            Label(self, text=attribute).pack(side=RIGHT)
        self.include = False

    def change(self):
        vars_list = []
        for var in self.vars:
            vars_list.append(var.get())
        # include was highlighted, exclude will now be chosen
        if vars_list[0] == 1 and self.include:
            self.vars[0].set(0)
            self.vars[1].set(1)
            self.include = False
        # exclude was highlighted, include will now be chosen
        elif vars_list[1] == 1 and not self.include:
            self.vars[0].set(1)
            self.vars[1].set(0)
            self.include = True


# similar to HTML_recon, this pulls up a tkinter frame with the list of duplicate schedules and their information
# so that the user can decide which to keep and which to delete
def duplicate_processing(schedule_unique_ind, schedule_file):
    # then check if there are duplicate schedules
    duplicates = []
    non_duplicates = []
    for x in schedule_unique_ind:
        if x not in non_duplicates:
            non_duplicates.append(x)
        else:
            duplicates.append(x)
    # create small object to hold lists so that it can be returned after frame is used to find schedules
    all_sched = All_schedules(schedule_unique_ind, duplicates)
    if len(duplicates) != 0:
        for dup in duplicates:
            print(dup, "is a duplicate schedule.")
        root = Tk()
        root.lift()
        root.title("Duplicate Schedules Editor")
        root.attributes('-topmost', True)
        root.after_idle(root.attributes, '-topmost', False)
        root.geometry('800x300+300+200')
        MainFrame(root, schedule_file, duplicates, schedule_unique_ind, all_sched).pack(side="top", fill="both", expand=True)
        root.mainloop()
    # returns either the edited list after choosing from duplicates or the untouched list if no duplicates
    return all_sched.duplicates, all_sched.schedule_unique_ind


def drop_down_question(options, freq):

    def quit_new_tab():
        master.quit()
        master.destroy()

    master = Tk()
    master.title('Choose Existing Freq')
    master.lift()
    master.attributes('-topmost', True)
    master.after_idle(master.attributes, '-topmost', False)

    column_var = StringVar()
    Label(master, text=str('Please Select Correct Base Frequency for: ' + freq)).pack(side=TOP)
    column_var.set("Correct Core Freq")

    column = OptionMenu(master, column_var, *options)
    column.pack(side=TOP)
    button = Button(master, text="Submit", command=quit_new_tab)
    button.pack(side=TOP)
    master.mainloop()

    return column_var.get()


def get_post_info(format, opt_sched_file, postal_facility, html_files, called_from):
    # "new" html format
    if format == 'new':
        # if called from opti app the HTML Check files may be different, so use the files loaded into the opti app
        if called_from == 'from_optimizer':
            html_files = [html_files]
        # if not called from optimizer use the htmls in html check
        else:
            html_files = []
            for file in os.listdir('HTML Check/'):
                if file.endswith('.htm') or file.endswith('.html'):
                    html_files.append(str('HTML Check/' + file))
        soup = BeautifulSoup(open(html_files[0]), "html.parser")
        main_table = soup.findAll('div')
        # removing all empty "rows" from the table which are just lines that look like <div></div>
        main_table = [row for row in main_table if row.getText() and row.getText() != []]
        postal_facility_final = main_table[0].findAll('tr')[1].findAll('td')[1].getText()
        row = lookups.loc[(lookups['Postal Facility Name'] == postal_facility_final)].index[0]
    # vital DB excel files, postal facility is given as input
    elif format == 'from db':
        postal_facility_final = postal_facility
        row = lookups.loc[(lookups['Postal Facility Name'] == postal_facility_final)].index[0]
    # old HTML files that use schedule summaries
    else:
        opt_sched_file = pd.read_excel(opt_sched_file, sheet_name='Schedule Summaries')
        short_name = opt_sched_file.loc[0, 'Site Name']
        row = lookups.loc[(lookups['Short Name'] == short_name)].index[0]
        postal_facility_final = lookups.loc[row, 'Postal Facility Name']

    pvs_names = [lookups.loc[row, 'PVS'], lookups.loc[row, 'PVS'], lookups.loc[row, 'Alternate PVS Name'],
                 lookups.loc[row, 'Alternate PVS Name 2'], lookups.loc[row, 'Alternate PVS Name 3'],
                 lookups.loc[row, 'Alternate PVS Name 4'], lookups.loc[row, 'Alternate PVS Name 5']]
    pdc_names = [lookups.loc[row, 'HCR P&DC Name'], lookups.loc[row, 'PVS P&DC'],
                 lookups.loc[row, 'Alternate P&DC Name'],
                 lookups.loc[row, 'Alternate P&DC Name 2'], lookups.loc[row, 'Alternate P&DC Name 3'],
                 lookups.loc[row, 'Alternate P&DC Name 4'], lookups.loc[row, 'Alternate P&DC Name 5']]
    pvs_and_pdc_names = pvs_names + pdc_names

    return pvs_names, pdc_names, pvs_and_pdc_names, postal_facility_final


# takes the schedules and stop excels from that are pulled from the database and creates
# schedule and stop objects with that information
def generate_schedules_db_format(schedule_file_df, stops_file_df, pvs_names, pdc_names):
    # converting schedule numbers to strings so that the unique identifier string can be created
    schedule_file_df['SCH_SCHED_NBR'] = [str(x) for x in schedule_file_df['SCH_SCHED_NBR'].tolist()]
    stops_file_df['SCH_SCHED_NBR'] = [str(x) for x in stops_file_df['SCH_SCHED_NBR'].tolist()]
    try:
        schedule_file_df['SCH_EFFECT_DTM'] = [dateparser.parse(str(x[:9])).strftime('%m/%d/%Y') for x in schedule_file_df['SCH_EFFECT_DTM'].tolist()]
        schedule_file_df['END_DT'] = [dateparser.parse(str(x[:9])).strftime('%m/%d/%Y') for x in schedule_file_df['END_DT'].tolist()]
    except:
        schedule_file_df['SCH_EFFECT_DTM'] = [dateparser.parse(str(x)[:10]).strftime('%m/%d/%Y') for x in
                                              schedule_file_df['SCH_EFFECT_DTM'].tolist()]
        schedule_file_df['END_DT'] = [dateparser.parse(str(x)[:10]).strftime('%m/%d/%Y') for x in
                                      schedule_file_df['END_DT'].tolist()]
    # creating unique identifier to differentiate between future and inservice schedules with the same sched #
    schedule_file_df['UNIQ_IND'] = schedule_file_df[['SCH_SCHED_NBR', 'DECODE_DESC', 'SITE_NAME']].agg('_'.join, axis=1)
    stops_file_df['UNIQ_IND'] = stops_file_df[['SCH_SCHED_NBR', 'DECODE_DESC', 'SITE_NAME']].agg('_'.join, axis=1)
    # first check if all schedule numbers that are in stops are also in the schedules file
    schedule_nums = list((schedule_file_df.loc[:, 'UNIQ_IND'].tolist()))
    stops_schedules_nums = list(set(stops_file_df.loc[:, 'UNIQ_IND'].tolist()))
    schedule_nums.sort()
    stops_schedules_nums.sort()
    if len(list(set(schedule_file_df['SCH_SCHED_NBR']))) != len(list(set(stops_file_df['SCH_SCHED_NBR']))):
        # puts any schedule number that is not in the stops file in the list
        sched_file_only = [schd for schd in list(set(schedule_file_df['SCH_SCHED_NBR'])) if schd not
                           in list(set(stops_file_df['SCH_SCHED_NBR']))]
        # puts any schedule number that is not in the schedules file in the list
        stop_file_only = [stop_sched for stop_sched in list(set(stops_file_df['SCH_SCHED_NBR'])) if stop_sched not
                          in list(set(schedule_file_df['SCH_SCHED_NBR']))]
        if len(sched_file_only) != 0:
            for schedule in sched_file_only:
                print("Schedule", schedule, "is in the schedules file and not the stops file. It won't be analyzed.")
                schedule_nums.remove(schedule)
        if len(stop_file_only) != 0:
            for schedule in stop_file_only:
                print("Schedule", schedule, "is in the stops file and not the schedules file. It won't be analyzed.")
                stops_schedules_nums.remove(schedule)
    # must do remove duplicates before the schedules are made
    duplicates, schedule_nums = duplicate_processing(schedule_nums, schedule_file_df)
    all_vital_schedules = []
    for schedule in schedule_nums:
        if schedule in stops_schedules_nums:
            schedule_df = schedule_file_df[(schedule_file_df['UNIQ_IND'] == schedule)]
            stops_df = stops_file_df[(stops_file_df['UNIQ_IND'] == schedule)]
            if schedule_df.empty or stops_df.empty:
                print('Schedule', schedule, 'DF is empty!')
            new_schedule = VITALSchedule_DB_or_NF("db format", schedule_df, stops_df, 'nf_series', 'nf_filename',
                                            pvs_names, pdc_names)
            all_vital_schedules.append(new_schedule)
    return all_vital_schedules


# takes HTML files with the new format in HTML Check folder and breaks
# them down into schedule objects which contain stop objects
def generate_schedules_nf(html_files, pvs_names, pdc_names, pvs_and_pdc_names):
    all_vital_schedules = []
    for file_name in html_files:
        soup = BeautifulSoup(open(file_name), "html.parser")
        main_table = soup.findAll('div')
        # removing all empty "rows" from the table which are just lines that look like <div></div>
        main_table = [row for row in main_table if row.getText() and row.getText() != []]
        text_in_rows = []
        for schedule in main_table:
            findalltr = schedule.findAll('tr')
            tr_iter = 0
            while tr_iter < len(findalltr):
                row_text = []
                for td in findalltr[tr_iter].findAll('td'):
                    if td.getText():
                        if '\n' in td.getText() or '\t' in td.getText() or '\xa0' in td.getText():
                            text = td.getText().replace('\n', '').replace('\t', '').replace('\xa0', '')
                            if 'ServicePoint' in text:
                                text = 'Service Point'
                        else:
                            text = td.getText()
                        row_text.append(text)
                text_in_rows.append(row_text)
                tr_iter += 1
        # creates a list of the start index of each schedule
        first_list = [x for x in list(range(len(text_in_rows))) if text_in_rows[x] and 'U.S.' in text_in_rows[x][0]]
        for start, sched in enumerate(first_list):
            if start != len(first_list)-1:
                schedule = text_in_rows[first_list[start]:first_list[start+1]]
            else:
                schedule = text_in_rows[first_list[start]:]
            newSchedule = VITALSchedule_DB_or_NF("new htmls", 'db_sched_df', 'db_stops_df', schedule,
                                           file_name.split('/')[-1], pvs_names, pdc_names)
            all_vital_schedules.append(newSchedule)
    return all_vital_schedules


# takes Schedules file from optimization application and creates schedule objects from the As Read Schedules df
def generate_schedules_of(schedules_file_df, pvs_names, pdc_names):
    schedule_nums = list(set((schedules_file_df.loc[:, 'Schedule Num'].tolist())))
    all_vital_schedules = []
    for schedule in schedule_nums:
        schedule_df = schedules_file_df[(schedules_file_df['Schedule Num'] == schedule)]
        new_schedule = VITALSchedule_OF(schedule_df, schedule_df.index, pvs_names, pdc_names)
        all_vital_schedules.append(new_schedule)
    return all_vital_schedules


class MainFrame(Frame):
    def __init__(self, root, schedules_file, duplicates, schedule_unique_ind, all_sched):
        Frame.__init__(self, root)
        self.scrollFrame = ScrollFrame(self)  # add a new scrollable frame.

        lines = []
        for dup in duplicates:
            rows = schedules_file[(schedules_file['SCH_SCHED_NBR'] == (dup.split('_')[0]))]
            indicies = rows.index.tolist()
            for row in indicies:
                line = Checkbar(self.scrollFrame.viewPort, [rows.loc[row, :], dup])
                line.pack(side=TOP, fill=X)
                line.config(relief=GROOVE, bd=10)
                lines.append(line)

        def quit():
            for line in lines:
                if not line.include:
                    schedule_unique_ind.remove(line.schedule_unique_ind)
                    all_sched.duplicates[all_sched.duplicates.index(line.schedule_unique_ind)] = \
                        str(line.schedule_unique_ind + '_' + line.effective_date + '_' + line.end_effective_date +
                            '_' + line.start_time)
                    print("Schedule", line.schedule_unique_ind, "with effective date", line.effective_date,
                          " will not be used.")
            all_sched.schedule_unique_ind = schedule_unique_ind
            root.quit()

        Button(root, text='Submit', command=quit).pack(side=BOTTOM)

        # when packing the scrollframe, we pack scrollFrame itself (NOT the viewPort)
        self.scrollFrame.pack(side="top", fill="both", expand=True)


# creating stop objects based on the vital database excel format and how its read in
class Stop_DB_or_NF:
    def __init__(self, type, db_index, db_schedule, db_stops_df, pvs_names, pdc_names, nf_series, nf_vital_schedule):
        if type == "db format":
            self.pvs_names = pvs_names
            self.pdc_names = pdc_names
            self.status = db_schedule.status
            self.schedule_paid_time = db_schedule.paid_time
            # self.schedule_paid_time = schedule.paid_time
            # try and except statements are for renamed columns in the data - hopefully taken out later
            try:
                self.arrive_time = datetime.strptime(str(db_stops_df.loc[db_index, 'ARR_TIME']), '%I:%M %p').replace(year=2022,
                                                                                                               month=1,
                                                                                                               day=1)
                self.depart_time = datetime.strptime(str(db_stops_df.loc[db_index, 'DEP_TIME']), '%I:%M %p').replace(year=2022,
                                                                                                               month=1,
                                                                                                               day=1)
            except:
                self.arrive_time = datetime.strptime(str(db_stops_df.loc[db_index, 'ARR_TIME'][:15] + ' ' +
                                                     db_stops_df.loc[db_index, 'ARR_TIME'][-2:]).replace('.', ':'),
                                                     '%d-%b-%y %I:%M %p').replace(year=2022, month=1, day=1)
                self.depart_time = datetime.strptime(str(db_stops_df.loc[db_index, 'DEP_TIME'][:15] + ' ' +
                                                         db_stops_df.loc[db_index, 'DEP_TIME'][-2:]).replace('.', ':'),
                                                     '%d-%b-%y %I:%M %p').replace(year=2022, month=1, day=1)
            self.duration = calculate_duration(self.depart_time, self.arrive_time, 'stop')
            self.stop_num = str(db_stops_df.loc[db_index, 'STOP_NBR'])
            try:
                self.stop_name = db_stops_df.loc[db_index, 'STOP_NAME']
            except:
                self.stop_name = db_stops_df.loc[db_index, 'FAC_NAME']
            try:
                self.site_id = db_stops_df.loc[db_index, 'PVS_SITE_ID']
            except:
                self.site_id = db_stops_df.loc[db_index, 'FAC_ID']
            self.spotter_stop = self.stop_name.lower() == 'spotter'
            self.freq_code = db_schedule.frequency
            self.begin_eff_date = db_schedule.effective_date
            self.end_eff_date = db_schedule.end_effective_date
            self.times_route = db_schedule.times_route
            self.times_trip = db_schedule.times_trip
            self.run_num = db_schedule.run_num
            self.tour = db_schedule.tour
            self.postalization_time = 0
            self.postalization_reason_code = ''
            self.postalization_stop, self.pvs_stop, self.pdc_stop, self.standby_stop = False, False, False, False
            self.lunch_stop = False
            self.lunch_reason = ''
            self.paid_time_reason = ''
            self.stop_info = [db_schedule.schedule_number, int(self.stop_num), str(self.arrive_time.strftime('%H:%M:%S')),
                              str(self.depart_time.strftime('%H:%M:%S')), self.stop_name, self.duration, '', '',
                              '', '', round(db_schedule.annual_trips, 2), '', self.freq_code,
                              db_schedule.tour, db_schedule.vehicle, db_schedule.effective_date,
                              db_schedule.end_effective_date, db_schedule.times_route, db_schedule.times_trip,
                              db_schedule.source_file]
            self.lunch_info = [db_schedule.schedule_number, int(self.stop_num), str(self.arrive_time.strftime('%H:%M:%S')),
                               str(self.depart_time.strftime('%H:%M:%S')), self.stop_name, self.duration, '', '',
                               '', '', self.schedule_paid_time, round(db_schedule.annual_trips, 2), '', self.freq_code,
                               db_schedule.tour, db_schedule.vehicle, db_schedule.effective_date,
                               db_schedule.end_effective_date, db_schedule.times_route, db_schedule.times_trip,
                               db_schedule.source_file]
            self.get_extra_info(db_schedule)
        elif type == "new htmls":
            self.pvs_names = pvs_names
            self.pdc_names = pdc_names
            self.stop_num = int(nf_series[0])
            time_objs = [item for item in nf_series if re.match("[0-9][0-9]:[0-9][0-9]", item)]
            time_index = nf_series.index(time_objs[0])
            self.arrive_time = dateparser.parse(nf_series[time_index]).replace(month=1, day=1, year=2021)
            self.depart_time = dateparser.parse(nf_series[time_index + 1]).replace(month=1, day=1, year=2021)
            self.stop_name = nf_series[time_index - 1]
            self.spotter_stop = self.stop_name.lower() == 'spotter'
            self.duration = calculate_duration(self.depart_time, self.arrive_time, 'stop')
            self.freq_code = nf_vital_schedule.frequency
            self.begin_eff_date = nf_vital_schedule.effective_date
            self.end_eff_date = nf_vital_schedule.end_effective_date
            self.times_route = nf_vital_schedule.times_route
            self.times_trip = nf_vital_schedule.times_trip
            self.run_num = nf_vital_schedule.run_num
            self.tour = nf_vital_schedule.tour
            self.postalization_time = 0
            self.postalization_reason_code = ''
            self.postalization_stop, self.pvs_stop, self.pdc_stop, self.standby_stop = False, False, False, False
            self.lunch_stop = False
            self.lunch_reason = ''
            self.schedule_paid_time = nf_vital_schedule.paid_time
            self.stop_info = [nf_vital_schedule.schedule_number, self.stop_num, str(self.arrive_time.strftime('%H:%M:%S')),
                              str(self.depart_time.strftime('%H:%M:%S')), self.stop_name, self.duration, '', '',
                              '', '', round(nf_vital_schedule.annual_trips, 2), '', self.freq_code,
                              nf_vital_schedule.tour, nf_vital_schedule.vehicle, nf_vital_schedule.effective_date,
                              nf_vital_schedule.end_effective_date, nf_vital_schedule.times_route, nf_vital_schedule.times_trip,
                              nf_vital_schedule.source_file]
            self.lunch_info = [nf_vital_schedule.schedule_number, self.stop_num,
                               str(self.arrive_time.strftime('%H:%M:%S')),
                               str(self.depart_time.strftime('%H:%M:%S')), self.stop_name, self.duration, '', '',
                               '', '', '', self.schedule_paid_time, round(nf_vital_schedule.annual_trips, 2), '',
                               self.freq_code,
                               nf_vital_schedule.tour, nf_vital_schedule.vehicle, nf_vital_schedule.effective_date,
                               nf_vital_schedule.end_effective_date, nf_vital_schedule.times_route, nf_vital_schedule.times_trip,
                               nf_vital_schedule.source_file]
            self.get_extra_info(nf_vital_schedule)

    def analyze_lunch(self, scenario):
        # paid time is 8 or greater than 8
        if scenario == 1 or scenario == 2:
            if self.duration < 30:
                self.lunch_reason = 'Lunch too short'
        # paid time is between 6 and 8 hours
        if scenario == 2:
            if self.duration > 30:
                self.lunch_reason = 'Lunch too long'

    def get_extra_info(self, schedule):
        # if stop name is a PVS or P&DC, label it as postalization and add postalization time
        if self.stop_name in self.pvs_names or self.stop_name in self.pdc_names:
            self.postalization_stop = True
            if self.stop_name in self.pvs_names:
                self.pvs_stop = True
                if self.duration > 15:
                    self.postalization_time = self.duration - 15
                else:
                    self.postalization_time = 0
            else:
                self.pdc_stop = True
                if self.duration > 30:
                    self.postalization_time = self.duration - 30
                else:
                    self.postalization_time = 0
            # adding postalization excess
            self.stop_info[6] = self.postalization_time
            # adding annual excess
            self.stop_info[11] = round(float(schedule.annual_trips * self.postalization_time), 2)
        if self.stop_name.lower() == 'standby time':
            self.standby_stop = True
            # adding annual excess
            self.stop_info[11] = round(float(schedule.annual_trips * self.duration), 2)
        elif self.stop_name.lower() == 'spotter':
            self.spotter_stop = True
            self.stop_info[14] = 'TTO'
        elif self.stop_name.lower() == 'lunch':
            self.lunch_stop = True
            if self.schedule_paid_time == 8:
                self.analyze_lunch(1)
                self.paid_time_reason = 'Normal'
            elif self.schedule_paid_time > 8:
                self.analyze_lunch(1)
                self.paid_time_reason = 'Greater than 8'
            elif 6 <= self.schedule_paid_time < 8:
                self.analyze_lunch(2)
                self.paid_time_reason = 'Between 6 and 8'
            elif self.schedule_paid_time < 6:
                self.analyze_lunch(2)
                self.paid_time_reason = 'Less than 6'
            # checking  if the lunch is too early or too late, all 3 scenarios need
            difference = (self.arrive_time - schedule.stops[0].arrive_time).total_seconds() / 60
            if difference < 0:
                difference = difference + 1440
            if difference / 60 > 6:
                if self.lunch_reason != '':
                    self.lunch_reason = self.lunch_reason + ', too late'
                else:
                    self.lunch_reason = 'Lunch too late'
            elif difference / 60 < 2:
                if self.lunch_reason != '':
                    self.lunch_reason = self.lunch_reason + ', too early'
                else:
                    self.lunch_reason = 'Lunch too early'
            elif self.lunch_reason == '':
                self.lunch_reason = 'Good lunch'
            # adding schedule start time to stop_info
            self.lunch_info[6] = str(schedule.stops[0].arrive_time.strftime('%H:%M:%S'))
            # adding paid time reason
            self.lunch_info[7] = self.paid_time_reason
            # adding lunch reason
            self.lunch_info[8] = self.lunch_reason


# creating a vital schedule object based on the vital database excel format
class VITALSchedule_DB_or_NF:
    def __init__(self, type, db_schedules_dataframe, db_stops_dataframe, nf_series, nf_filename, pvs_names, pdc_names):
        self.lunch_stop, self.first_stop, self.second_stop, self.before_last_stop, self.last_stop, self.start_time, \
        self.end_time, self.pvs_to_pdc_stop, self.pdc_to_pvs_stop = None, None, None, None, None, None, None, None, \
                                                                    None
        self.stops, self.problem_stops, self.lunch_stops = [], [], []
        self.total_pvs_duration, self.total_pdc_duration, self.spotter_time, self.standby_time = 0, 0, 0, 0
        self.has_lunch, self.spotter_schedule = False, False
        self.pvs_names = pvs_names
        self.pdc_names = pdc_names
        if type == "db format":
            row = db_schedules_dataframe.index[0]
            index = list(db_stops_dataframe.index)
            self.schedule_number = str(db_schedules_dataframe.loc[row, 'SCH_SCHED_NBR'])
            self.source_file = ''
            self.tour = str(db_schedules_dataframe.loc[row, 'TOUR_NBR'])
            self.run_num = str(db_schedules_dataframe.loc[row, 'RUN_NBR'])
            if db_schedules_dataframe.loc[row, 'TRACTOR_IND'] == 'Y':
                self.vehicle = 'TTO'
            else:
                self.vehicle = 'MVO'
            self.frequency = db_schedules_dataframe.loc[row, 'FRQ_CD']
            try:
                self.annual_trips = frequencies.loc[frequencies[(frequencies['Freq'] == self.frequency)].index[0],
                                                'Trips']
            except:
                print('Frequency code not recognized in PVS Tools and Templates/Frequency Codes for Zero Bases.xlsx')
                print('Freq: ', self.frequency)
                frequency_list = sorted([str(x)[:10] for x in list(set(frequencies['Core Freq']))])
                frequency_list_drop_down = []
                for code in frequency_list:
                    description = frequencies.loc[frequencies[(frequencies['Core Freq'] == code)].index[0], 'Description']
                    frequency_list_drop_down.append(str(code + ': ' + description))
                new_freq = drop_down_question(frequency_list_drop_down, self.frequency)
                new_freq = new_freq.split(':')[0]
                self.annual_trips = frequencies.loc[frequencies[(frequencies['Core Freq'] == new_freq)].index[0],
                                                    'Trips']
            test_duration = datetime.strptime(db_schedules_dataframe.loc[row, 'SCH_DURATION'], '%H:%M').time()
            try:
                min_duration = float(int(str(db_schedules_dataframe.loc[row, 'SCH_DURATION']).replace(':', '.')[2:]) / 60)
            except:
                min_duration = float(int(str(db_schedules_dataframe.loc[row, 'SCH_DURATION']).replace(':', '.')[3:]) / 60)
            if len(str(db_schedules_dataframe.loc[row, 'SCH_DURATION'])) < 5:
                self.duration = float(str(db_schedules_dataframe.loc[row, 'SCH_DURATION'])[0:1]) + min_duration
            else:
                self.duration = float(str(db_schedules_dataframe.loc[row, 'SCH_DURATION'])[1:2]) + min_duration
            self.annual_hours = float(self.duration * self.annual_trips)
            self.annual_miles = float(db_schedules_dataframe.loc[row, 'MILEAGE_NBR'] * self.annual_trips)
            self.daily_miles = db_schedules_dataframe.loc[row, 'MILEAGE_NBR']
            self.effective_date = db_schedules_dataframe.loc[row, 'SCH_EFFECT_DTM']
            self.end_effective_date = db_schedules_dataframe.loc[row, 'END_DT']
            if self.end_effective_date[-4:] == '1999':
                self.end_effective_date = self.end_effective_date[:-4] + '2999'
            self.times_trip = db_schedules_dataframe.loc[row, 'TIME_TRIP_ID']
            self.times_route = db_schedules_dataframe.loc[row, 'SCH_ROUTE_ID']
            if 'DECODE_DESC' in db_schedules_dataframe.columns.tolist():
                self.status = db_schedules_dataframe.loc[row, 'DECODE_DESC']
            else:
                self.status = 'not_listed'
            if 'AREA_NAME' in db_schedules_dataframe.columns.tolist():
                self.area = db_schedules_dataframe.loc[row, 'AREA_NAME']
            else:
                self.area = 'N/A'
            try:
                self.number_of_pvs_stops = len([x for x in list(range(len(index))) if index[x] and
                                            db_stops_dataframe.loc[index[x], 'STOP_NAME'] in pvs_names])
                self.number_of_pdc_stops = len([x for x in list(range(len(index))) if index[x] and
                                            db_stops_dataframe.loc[index[x], 'STOP_NAME'] in pdc_names])
            except:
                self.number_of_pvs_stops = len([x for x in list(range(len(index))) if index[x] and
                                                db_stops_dataframe.loc[index[x], 'FAC_NAME'] in pvs_names])
                self.number_of_pdc_stops = len([x for x in list(range(len(index))) if index[x] and
                                                db_stops_dataframe.loc[index[x], 'FAC_NAME'] in pdc_names])
            self.number_of_stops = db_schedules_dataframe.loc[row, 'TOT_STOP_CNT']
            # have to get the lunch duration before stop objects are created so that they can hold sched_paid_time variable
            lunch_stop_index = [index[x] for x in list(range(len(index))) if index[x] and
                                db_stops_dataframe.loc[index[x], 'STOP_NAME'] in ['Lunch', 'LUNCH', 'lunch']]
            if len(lunch_stop_index) > 0:
                lunch_index = lunch_stop_index[0]
                lunch_duration = calculate_duration(datetime.strptime(str(db_stops_dataframe.loc[lunch_index, 'DEP_TIME']),
                                                                      '%I:%M %p').replace(year=2022, month=1, day=1),
                                                    datetime.strptime(str(db_stops_dataframe.loc[lunch_index, 'ARR_TIME']),
                                                                      '%I:%M %p').replace(year=2022, month=1, day=1),
                                                    'stop')
                self.paid_time = round(float(self.duration - (lunch_duration/60)), 2)
            else:
                self.paid_time = self.duration
            self.get_stops("db format", db_stops_dataframe, index, 'nf_series')
        elif type == "new htmls":
            self.schedule_number = nf_series[1][-1]
            self.source_file = nf_filename
            self.tour = nf_series[1][3]
            self.run_num = nf_series[1][4]
            if len(nf_series[1]) == 9:
                self.vehicle = 'TTO'
                freq_index = -3
            else:
                self.vehicle = 'MVO'
                freq_index = -2
            self.frequency = str(nf_series[1][freq_index])
            self.annual_hours = nf_series[7][3].replace(':', '.')
            self.annual_miles = nf_series[7][2].replace(':', '.')
            self.annual_trips = frequencies[(frequencies['Freq'] == nf_series[1][freq_index])]['Trips'].tolist()[0]
            self.effective_date = dateparser.parse(nf_series[3][0]).strftime('%m/%d/%Y')
            self.end_effective_date = dateparser.parse(nf_series[3][1]).strftime('%m/%d/%Y')
            self.times_trip = nf_series[7][1]
            self.times_route = nf_series[7][0]
            self.duration = calculate_duration(dateparser.parse(nf_series[11][2]).replace(month=1, day=1, year=2021),
                                               dateparser.parse(nf_series[-1][3]).replace(month=1, day=1, year=2021),
                                               'sched')
            self.number_of_stops = len(nf_series[11:])
            self.number_of_pvs_stops = 0
            self.number_of_pdc_stops = 0
            for stop in nf_series[11:]:
                time_index = stop.index([item for item in stop if re.match("[0-9][0-9]:[0-9][0-9]", item)][0])
                if stop[time_index-1] in self.pdc_names:
                    self.number_of_pdc_stops += 1
                elif stop[time_index-1] in self.pvs_names:
                    self.number_of_pvs_stops += 1
            standby_duration = 0
            spotter_duration = 0
            for stop in nf_series[11:]:
                if stop[1].lower() == 'standby time':
                    standby_duration = standby_duration + calculate_duration(
                        dateparser.parse(stop[3]).replace(month=1, day=1, year=2021),
                        dateparser.parse(stop[2]).replace(month=1, day=1, year=2021), 'stop')
                elif stop[1].lower() == 'spotter':
                    spotter_duration = spotter_duration + calculate_duration(
                        dateparser.parse(stop[3]).replace(month=1, day=1, year=2021),
                        dateparser.parse(stop[2]).replace(month=1, day=1, year=2021), 'stop')
            self.standby_time = standby_duration
            self.spotter_time = spotter_duration
            lunch_duration = 0
            for stop in nf_series[11:]:
                if stop[1].lower() == 'lunch':
                    lunch_duration = calculate_duration(dateparser.parse(stop[3]).replace(month=1, day=1, year=2021),
                                                        dateparser.parse(stop[2]).replace(month=1, day=1, year=2021),
                                                        'stop')
            if lunch_duration != 0:
                self.paid_time = self.duration - lunch_duration / 60
            else:
                self.paid_time = self.duration
            self.get_stops("new htmls", 'db_stops_dataframe', 'db_index', nf_series)
        if not self.spotter_schedule:
            self.postalization_check()

        self.check_lunches()

    def get_stops(self, type, db_stops_dataframe, index, nf_series):
        if type == "db format":
            for stop_index in index:
                newStop = Stop_DB_or_NF("db format", stop_index, self, db_stops_dataframe, self.pvs_names,
                                        self.pdc_names, 'nf_series', 'nf_vital_schedule')
                # calculating number of postalization stops and duration
                if newStop.stop_name in self.pvs_names:
                    self.number_of_pvs_stops = self.number_of_pvs_stops + 1
                    self.total_pvs_duration = self.total_pvs_duration + newStop.duration
                elif newStop.stop_name in self.pdc_names:
                    self.number_of_pdc_stops = self.number_of_pdc_stops + 1
                    self.total_pdc_duration = self.total_pdc_duration + newStop.duration
                # getting lunch information and before stop info
                if newStop.stop_name.lower() == 'lunch':
                    newStop.lunch_stop = True
                    self.lunch_stop = newStop
                    self.has_lunch = True
                    if newStop.lunch_reason != 'Good lunch':
                        self.problem_stops.append(newStop)
                    self.stops[index.index(stop_index) - 1].before_lunch_stop = True
                    self.stops[index.index(stop_index) - 1].lunch_info[7] = newStop.paid_time_reason
                # checking if the last stop was lunch so that after lunch stop can be identified
                # only enters if it is not the very first or very last stop
                if index.index(stop_index) != len(index) - 1 and index.index(stop_index) != 0:
                    if self.stops[index.index(stop_index) - 1].lunch_stop:
                        newStop.after_lunch_stop = True
                        newStop.lunch_info[7] = self.stops[index.index(stop_index) - 1].paid_time_reason
                self.stops.append(newStop)
        elif type == "new htmls":
            stops = nf_series[11:]
            for stop in stops:
                newStop = Stop_DB_or_NF("new htmls", 'db_index', 'db_schedule', 'db_stops_df',
                                        self.pvs_names, self.pdc_names, stop, self)
                # calculating number of postalization stops and duration
                if newStop.stop_name in self.pvs_names:
                    self.number_of_pvs_stops = self.number_of_pvs_stops + 1
                    self.total_pvs_duration = self.total_pvs_duration + newStop.duration
                elif newStop.stop_name in self.pdc_names:
                    self.number_of_pdc_stops = self.number_of_pdc_stops + 1
                    self.total_pdc_duration = self.total_pdc_duration + newStop.duration
                # getting lunch information and before stop info
                if newStop.stop_name.lower() == 'lunch':
                    newStop.lunch_stop = True
                    self.lunch_stop = newStop
                    self.has_lunch = True
                    if newStop.lunch_reason != 'Good lunch':
                        self.problem_stops.append(newStop)
                    self.stops[stops.index(stop) - 1].before_lunch_stop = True
                # checking if the last stop was lunch so that after lunch stop can be identified
                # only enters if it is not the very first or very last stop
                if stops.index(stop) != len(stops) - 1 and stops.index(stop) != 0:
                    if self.stops[stops.index(stop) - 1].lunch_stop:
                        newStop.after_lunch_stop = True
                self.stops.append(newStop)

    def find_postalization_stops(self):
        stops_to_remove = []
        for stop in self.stops:
            # changes stop name to all lower case so that case is ignored when analyzing
            if stop.stop_name.lower() in standby_list:
                if stop.stop_name.lower() == 'standby time':
                    self.standby_time = self.standby_time + stop.duration
                if stop.stop_name.lower() == 'spotter':
                    self.spotter_time = self.spotter_time + stop.duration
                stops_to_remove.append(stop)

        # initializes list and puts all original stops into so as not to mess them up
        postalization_stops = []
        for stop in self.stops:
            postalization_stops.append(stop)
        # removes all stops which have a name that is not an actual stop or PVS/P&DC
        for stop in stops_to_remove:
            postalization_stops.remove(stop)
        # sets first, last, second, second to last stops to check if it stops at PDC and PVS
        self.first_stop = postalization_stops[0]
        self.last_stop = postalization_stops[-1]
        self.second_stop = postalization_stops[1]
        self.before_last_stop = postalization_stops[-2]

        spotter_stops = [x for x in self.stops if x.stop_name.lower() == 'spotter']
        if len(self.stops) == 5 and len(spotter_stops) == 2:
            self.spotter_schedule = True

        # if len(postalization_stops) <= 4:
        #     if (self.stops[1].stop_name.lower() == 'spotter' or self.stops[2].stop_name.lower() == 'spotter') \
        #             and (
        #             self.stops[-1].stop_name.lower() == 'spotter' or self.stops[-2].stop_name.lower() == 'spotter'):
        #         self.spotter_schedule = True

    def postalization_check(self):
        self.find_postalization_stops()
        self.start_time = self.stops[0].arrive_time
        self.end_time = self.stops[-1].depart_time

        # first and second stop check
        # checks if first stop is PVS/P&DC and second isn't
        if self.first_stop.postalization_stop and not self.second_stop.postalization_stop:
            self.first_stop.postalization_time = self.first_stop.duration - 45
            if self.first_stop.pvs_stop:
                self.first_stop.postalization_reason_code = 'Does not stop at P&DC'
            else:
                self.first_stop.postalization_reason_code = 'Does not stop at PVS'
            self.first_stop.stop_info[6] = self.first_stop.postalization_time
            self.first_stop.stop_info[7] = self.first_stop.postalization_reason_code
            if self.first_stop.postalization_time > 0:
                self.first_stop.stop_info[11] = round(float(self.annual_trips*self.first_stop.postalization_time), 2)
            self.problem_stops.append(self.first_stop)
        # checks if second stop is PVS/P&DC and first isn't
        elif self.second_stop.postalization_stop and not self.first_stop.postalization_stop:
            self.second_stop.postalization_time = self.second_stop.duration - 45
            if self.second_stop.pvs_stop:
                self.second_stop.postalization_reason_code = 'Does not stop at P&DC'
            else:
                self.second_stop.postalization_reason_code = 'Does not stop at PVS'
            self.second_stop.stop_info[6] = self.second_stop.postalization_time
            self.second_stop.stop_info[7] = self.second_stop.postalization_reason_code
            if self.second_stop.postalization_time > 0:
                self.second_stop.stop_info[11] = round(float(self.annual_trips*self.second_stop.postalization_time), 2)
            self.problem_stops.append(self.second_stop)
        # else if both are neither PVS/P&DC
        elif not self.first_stop.postalization_stop and not self.second_stop.postalization_stop:
            self.first_stop.postalization_time = self.first_stop.duration - 45
            self.first_stop.postalization_reason_code = 'Did not start at PVS or P&DC - incorrect postalization'
            self.first_stop.stop_info[6] = self.first_stop.postalization_time
            self.first_stop.stop_info[7] = self.first_stop.postalization_reason_code
            if self.first_stop.postalization_time > 0:
                self.first_stop.stop_info[11] = round(float(self.annual_trips*self.first_stop.postalization_time), 2)
            self.problem_stops.append(self.first_stop)

        # before last and last stop check
        # checks if last stop is PVS/P&DC and before last isn't
        if self.last_stop.postalization_stop and not self.before_last_stop.postalization_stop:
            self.last_stop.postalization_time = self.last_stop.duration - 45
            if self.last_stop.pvs_stop:
                self.last_stop.postalization_reason_code = 'Does not stop at P&DC'
            else:
                self.last_stop.postalization_reason_code = 'Does not stop at PVS'
            self.last_stop.stop_info[6] = self.last_stop.postalization_time
            self.last_stop.stop_info[7] = self.last_stop.postalization_reason_code
            if self.last_stop.postalization_time > 0:
                self.last_stop.stop_info[11] = round(float(self.annual_trips*self.last_stop.postalization_time), 2)
            self.problem_stops.append(self.last_stop)
        # checks if before last stop is PVS/P&DC and last isn't
        elif not self.last_stop.postalization_stop and self.before_last_stop.postalization_stop:
            self.before_last_stop.postalization_time = self.before_last_stop.duration - 45
            if self.before_last_stop.pvs_stop:
                self.before_last_stop.postalization_reason_code = 'Does not stop at P&DC'
            else:
                self.before_last_stop.postalization_reason_code = 'Does not stop at PVS'
            self.before_last_stop.stop_info[6] = self.before_last_stop.postalization_time
            self.before_last_stop.stop_info[7] = self.before_last_stop.postalization_reason_code
            if self.before_last_stop.postalization_time > 0:
                self.before_last_stop.stop_info[11] = round(float(self.annual_trips*self.before_last_stop.postalization_time), 2)
            self.problem_stops.append(self.before_last_stop)
        # else if both are neither PVS/P&DC
        elif not self.last_stop.postalization_stop and not self.before_last_stop.postalization_stop:
            self.last_stop.postalization_time = self.last_stop.duration - 45
            self.last_stop.postalization_reason_code = 'Did not end at PVS or P&DC - incorrect postalization'
            self.last_stop.stop_info[6] = self.last_stop.postalization_time
            self.last_stop.stop_info[7] = self.last_stop.postalization_reason_code
            if self.last_stop.postalization_time > 0:
                self.last_stop.stop_info[11] = round(float(self.annual_trips*self.last_stop.postalization_time), 2)
            self.problem_stops.append(self.last_stop)

        if self.first_stop.postalization_reason_code == '' and self.second_stop.postalization_reason_code == '' \
                and self.stops[1].stop_name.lower() not in standby_list:
            pvs_to_pdc_duration = calculate_duration(self.second_stop.arrive_time, self.first_stop.depart_time,
                                                     'stop')
            pvs_to_pdc_stop_nums = str(str(self.first_stop.stop_num) + ' - ' + str(self.second_stop.stop_num))
            if pvs_to_pdc_duration > 5:
                self.pvs_to_pdc_stop = [[self.schedule_number, pvs_to_pdc_stop_nums, 'PVS -> P&DC',
                                         str(self.first_stop.depart_time.strftime('%H:%M:%S')),
                                         str(self.second_stop.arrive_time.strftime('%H:%M:%S')), pvs_to_pdc_duration,
                                         self.annual_trips, round(float(self.annual_trips * pvs_to_pdc_duration), 2),
                                         self.frequency, self.tour, self.vehicle, self.source_file], []]
                self.pvs_to_pdc_stop[1] = [self.schedule_number, pvs_to_pdc_stop_nums,
                                           str(self.first_stop.depart_time.strftime('%H:%M:%S')),
                                           str(self.second_stop.arrive_time.strftime('%H:%M:%S')), 'PVS -> P&DC',
                                           pvs_to_pdc_duration, '', '', '', '', self.annual_trips,
                                           round(float(self.annual_trips * pvs_to_pdc_duration), 2), self.frequency,
                                           self.tour, self.vehicle, self.effective_date, self.end_effective_date,
                                           self.times_route, self.times_trip, self.source_file]
                if pvs_to_pdc_duration > 5:
                    self.problem_stops.append(self.pvs_to_pdc_stop[1])
        if self.last_stop.postalization_reason_code == '' and self.before_last_stop.postalization_reason_code == '' \
                and self.stops[-2].stop_name.lower() not in standby_list:
            pdc_to_pvs_duration = calculate_duration(self.last_stop.arrive_time, self.before_last_stop.depart_time,
                                                     'stop')
            pdc_to_pvs_stop_nums = str(str(self.before_last_stop.stop_num) + ' - ' + str(self.last_stop.stop_num))
            if pdc_to_pvs_duration > 5:
                self.pdc_to_pvs_stop = [[self.schedule_number, pdc_to_pvs_stop_nums, 'P&DC -> PVS',
                                         str(self.before_last_stop.depart_time.strftime('%H:%M:%S')),
                                         str(self.last_stop.arrive_time.strftime('%H:%M:%S')), pdc_to_pvs_duration,
                                         self.annual_trips, round(float(self.annual_trips * pdc_to_pvs_duration), 2),
                                         self.frequency, self.tour, self.vehicle, self.source_file], []]
                self.pdc_to_pvs_stop[1] = [self.schedule_number, pdc_to_pvs_stop_nums,
                                           str(self.before_last_stop.depart_time.strftime('%H:%M:%S')),
                                           str(self.last_stop.arrive_time.strftime('%H:%M:%S')), 'P&DC -> PVS',
                                           pdc_to_pvs_duration, '', '', '', '', self.annual_trips,
                                           round(float(self.annual_trips * pdc_to_pvs_duration), 2), self.frequency,
                                           self.tour, self.vehicle, self.effective_date, self.end_effective_date,
                                           self.times_route, self.times_trip, self.source_file]
                if pdc_to_pvs_duration > 5:
                    self.problem_stops.append(self.pdc_to_pvs_stop[1])

    def check_lunches(self):
        if self.has_lunch:
            index = self.stops.index(self.lunch_stop)
            before_lunch_stop = self.stops[index - 1]
            if before_lunch_stop.stop_name.lower() == 'standby time' or\
                    before_lunch_stop.stop_name.lower() == 'spotter':
                before_lunch_stop = self.stops[index - 2]
            if index + 1 >= len(self.stops):
                print('Schedule has lunch at the end of the day! ', self.schedule_number)
                return
            else:
                after_lunch_stop = self.stops[index + 1]
                if after_lunch_stop.stop_name.lower() == 'standby time' or \
                        after_lunch_stop.stop_name.lower() == 'spotter':
                    try:
                        after_lunch_stop = self.stops[index + 2]
                    except:
                        if self.stops.index(after_lunch_stop) == len(self.stops) - 1:
                            print("Schedule ", self.schedule_number, " has a spotter stop at the end of the schedule. "
                                                                     "Let the site know!")
                            after_lunch_stop.stop_info[8] = 'Lunch -> Spotter is end of schedule.'
                            self.problem_stops.append(after_lunch_stop)

            if before_lunch_stop and after_lunch_stop:
                # if lunch is the second to last stop followed by a spotter stop, exit the if statment
                if after_lunch_stop.stop_info[8] == 'Lunch -> Spotter is end of schedule.':
                    return
                # if different stop name before and after lunch
                if before_lunch_stop.stop_name != after_lunch_stop.stop_name:
                    # if one or both are PVS/PDC or standby they are fine, so check if they aren't one of those
                    if not ((before_lunch_stop.stop_name.lower() in standby_list) or
                            (before_lunch_stop.stop_name in self.pvs_names) or
                            (before_lunch_stop.stop_name in self.pdc_names)) and \
                            ((after_lunch_stop.stop_name.lower() in standby_list) or
                             (after_lunch_stop.stop_name in self.pvs_names) or
                             (after_lunch_stop.stop_name in self.pdc_names)):
                        before_lunch_stop.lunch_info[6] = self.start_time.strftime('%H:%M:%S')
                        before_lunch_stop.lunch_info[7] = self.lunch_stop.paid_time_reason
                        before_lunch_stop.lunch_info[8] = 'Different from after lunch stop'
                        after_lunch_stop.lunch_info[6] = self.start_time.strftime('%H:%M:%S')
                        after_lunch_stop.lunch_info[7] = self.lunch_stop.paid_time_reason
                        after_lunch_stop.lunch_info[8] = 'Different from before lunch stop'
                        self.problem_stops.append(before_lunch_stop)
                        self.problem_stops.append(after_lunch_stop)
                        self.lunch_stops.append(before_lunch_stop)
                        self.lunch_stops.append(self.lunch_stop)
                        self.lunch_stops.append(after_lunch_stop)
                        if self.lunch_stop.lunch_info[8] != 'Good lunch':
                            self.lunch_stop.lunch_info[8] = self.lunch_stop.lunch_info[8] + ', before/after different stops'
                        else:
                            self.lunch_stop.lunch_info[8] = 'Before/after different stops'
                        if self.lunch_stop not in self.problem_stops:
                            self.problem_stops.append(self.lunch_stop)

# ************************
# Scrollable Frame Class
# ************************
class ScrollFrame(Frame):
    def __init__(self, parent):
        super().__init__(parent)  # create a frame (self)

        self.canvas = Canvas(self, borderwidth=0, background="#ffffff")  # place canvas on self
        # place a frame on the canvas, this frame will hold the child widgets
        self.viewPort = Frame(self.canvas, background="#ffffff")
        self.vsb = Scrollbar(self, orient="vertical", command=self.canvas.yview)  # place a scrollbar on self
        self.canvas.configure(yscrollcommand=self.vsb.set)  # attach scrollbar action to scroll of canvas

        self.vsb.pack(side="right", fill="y")  # pack scrollbar to right of self
        self.canvas.pack(side="left", fill="both", expand=True)  # pack canvas to left of self and expand to fil
        self.canvas_window = self.canvas.create_window((4, 4), window=self.viewPort, anchor="nw",
                                                       # add view port frame to canvas
                                                       tags="self.viewPort")
        # bind an event whenever the size of the viewPort frame changes.
        self.viewPort.bind("<Configure>", self.onFrameConfigure)
        self.canvas.bind("<Configure>", self.onCanvasConfigure)

        # perform an initial stretch on render, otherwise the scroll region has a tiny border until the first resize
        self.onFrameConfigure(None)

    def onFrameConfigure(self, event):
        '''Reset the scroll region to encompass the inner frame'''
        self.canvas.configure(scrollregion=self.canvas.bbox(
            "all"))  # whenever the size of the frame changes, alter the scroll region respectively.

    def onCanvasConfigure(self, event):
        '''Reset the canvas window to encompass inner frame when required'''
        canvas_width = event.width
        # whenever the size of the canvas changes alter the window region respectively.
        self.canvas.itemconfig(self.canvas_window, width=canvas_width)