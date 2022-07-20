from GeneralMethods import ScheduleQuery, StopQuery, query_vehicles, query_rot, get_times_sched_for_rot
import pandas as pd
import time
from connect_to_db import get_connection
from Diagnostic import print_diagnostic
import os

type_dict = {'In-Service': 'I', 'Draft': 'D', 'Future': 'F', 'Discontinued': 'E'}

finance_numbers_excel = pd.read_excel("Updated finance numbers.xlsx", converters={"NEW Finance Number": str})
finance_numbers = ', '.join([str("'" + str(nbr) + "'") for nbr in finance_numbers_excel["NEW Finance Number"].tolist()])


def get_pvs_sites(conn):
    pvs_site_query = """select f.fac_id, p.pvs_site_cd, p.finance_nbr, 
    f.nass_cd,f.FAC_NAME, f.FAC_TYPE_CD, f.AREA_ID, f.AREA_NAME,F.FAC_STATUS
    from vitalprod.facility_t f 
    left join VITALPROD.PVS_SITE_T p  on p.pvs_site_id = f.fac_id --bring in finance nbr
    where 
    f.pvs_site_ind = 'Y'  --pvs sites indicator
    and p.finance_nbr in (""" + finance_numbers +  """) 
    --and p.finance_nbr <> '999999' --backup option for log issue 
    --f.fac_status = 'A' --active sites
    order by F.FAC_NAME, f.AREA_NAME, 
    f.fac_type_cd
    """
    data_schedule = pd.read_sql(pvs_site_query, conn)
    return data_schedule['FAC_NAME'].tolist()


def get_times_schedule_df(schedule_df, conn):
    schedule_df['TEMP_DATE'] = [x.strftime("%d-%b-%y").upper() for x in schedule_df['SCH_EFFECT_DTM']]
    schedules_dict = {schedule_df.loc[x, 'SCH_SCHED_NBR']: [schedule_df.loc[x, 'SCH_ROUTE_ID'], schedule_df.loc[x, 'TIME_TRIP_ID'],
                                                            schedule_df.loc[x, 'TEMP_DATE'], schedule_df.loc[x, 'SITE_NAME']] for x in
                      range(len(schedule_df))}
    all_times_sched_for_rot = []
    for schedule in schedules_dict.keys():
        times_sched_df = get_times_sched_for_rot(schedules_dict[schedule], conn)
        times_sched_df['SCH_SCHED_NBR'] = str(schedule)
        all_times_sched_for_rot.append(times_sched_df)
    print(list(schedule_df['SITE_NAME'])[0])
    final_df = pd.concat(all_times_sched_for_rot)
    return final_df


def print_zb_files(file_type, list_of_sites, separate, column_title, df):
    file_type_sites = list(set(df[column_title]))
    site_not_included = ''
    if len(file_type_sites) < len(list_of_sites):
        for file_site in list_of_sites:
            if file_site not in file_type_sites:
                site_not_included = site_not_included + ', ' + file_site
    if site_not_included != '':
        print(site_not_included, ' did not have any output for the ', file_type, 'file.')
        new_data_df = df[(df[column_title] == file_site)]
        new_data_df.to_excel('DataFiles\\' + file_site + ' ' + file_type.upper() + ' ' + time.strftime("%Y%m%d") +
                             '.xlsx', index=False, encoding='Cp1252')
    if separate:
        for site in file_type_sites:
            new_data_df = df[(df[column_title] == site)]
            new_data_df.to_excel('DataFiles\\' + site + ' ' + file_type.upper() + ' ' + time.strftime("%Y%m%d") +
                                 '.xlsx', index=False, encoding='Cp1252')
    elif not separate:
        df.to_excel('DataFiles\\' + '_'.join(file_type_sites[:2]) + ' ' + file_type.upper() + ' ' +
                    time.strftime("%Y%m%d") + '.xlsx', index=False, encoding='Cp1252')


def run_zb_files(list_of_sites, type, date, separate, conn):
    if type == 'Schedule Type':
        type = 'I'
    else:
        type = type_dict[type]
    if date == 'Date':
        date = ''
    newQuery = ScheduleQuery()
    # p is used for vehicle and ROT queries as well
    schedules_query, p = newQuery.build_reg_query(list_of_sites, type, date, [])
    # sites string is for vehicles and ROT queries later
    sites_string = ', '.join([str(':' + k) for k in p.keys()])
    data_schedule = pd.read_sql(schedules_query, conn, params=p)
    if data_schedule.empty:
        print("Oh no! Schedules file for ", list_of_sites, " is empty!")

    # getting times route file for ROT based on schedule info
    data_times_for_rot = get_times_schedule_df(data_schedule, conn)

    # prints the times route for ROT by site or together
    print_zb_files('Times_Route', list_of_sites, separate, 'SITE_NAME', data_times_for_rot)

    # prints query by site or together
    print_zb_files('Schedules', list_of_sites, separate, 'SITE_NAME', data_schedule)
    site_list = set(list(data_schedule['SITE_NAME']))
    # data_schedule['SITE_NAME'].tolist()
    for site in site_list:
        new_data_sched = data_schedule[(data_schedule['SITE_NAME'] == site)]
        duplicate_schedules = new_data_sched[new_data_sched.duplicated(subset=['SCH_SCHED_NBR'])]
        if not duplicate_schedules.empty:
            print("Duplicate schedules in ", site, ". Duplicate schedules are:")
            print(duplicate_schedules['SCH_SCHED_NBR'])

    newStopQuery = StopQuery()
    stops_query, params_stops = newStopQuery.build_reg_stop_query(list_of_sites, type, date, [])

    data_stops = pd.read_sql(stops_query, conn, params=p)
    if data_stops.empty:
        print("Oh no! Stops file for ", list_of_sites, " is empty!")
    print_zb_files('Stops', list_of_sites, separate, 'SITE_NAME', data_stops)

    vehicles_query = query_vehicles.replace('SITES_STRING_TO_REPLACE', sites_string)
    data_vehicles = pd.read_sql(vehicles_query, conn, params=p)
    if data_vehicles.empty:
        print("Oh no! Vehicles file for ", list_of_sites, " is empty!")
    print_zb_files('Vehicles', list_of_sites, separate, 'FAC_NAME', data_vehicles)

    rot_query = query_rot.replace('SITES_STRING_TO_REPLACE', sites_string)
    data_rot = pd.read_sql(rot_query, conn, params=p)
    if data_rot.empty:
        print("Oh no! ROT file for ", list_of_sites, " is empty!")
    print_zb_files('ROT', list_of_sites, separate, 'PVS_NAME', data_rot)


def process_duplicates(data_schedule, list_of_sites, conn, p):
    for site in list(set(data_schedule['SITE_NAME'])):
        new_data_sched = data_schedule[(data_schedule['SITE_NAME'] == site)]
        duplicate_schedules = new_data_sched[new_data_sched.duplicated()]
        if not duplicate_schedules.empty:
            data_schedule = data_schedule.drop_duplicates()

        for schedule in new_data_sched['SCH_SCHED_NBR'].tolist():
            if new_data_sched['SCH_SCHED_NBR'].tolist().count(schedule) > 1:
                print('Duplicate schedules for ', schedule)

# def run_zb_together_diff_types(list_of_sites, type, date, separate, conn):
#     if type == 'Schedule Type':
#         type = 'I'
#     else:
#         type = type_dict[type]
#     if date == 'Date':
#         date = ''
#     newQuery = ScheduleQuery()
#     # p is used for vehicle and ROT queries as well
#     schedules_query, p = newQuery.build_reg_query(list_of_sites, type, date, [])
#     # sites string is for vehicles and ROT queries later
#     sites_string = ', '.join([str(':' + k) for k in p.keys()])
#     data_schedule = pd.read_sql(schedules_query, conn, params=p)
#     if data_schedule.empty:
#         print("Oh no! Schedules file for ", list_of_sites, " is empty!")
#     # prints query by site or together
#     print_zb_files('Schedules', list_of_sites, separate, 'SITE_NAME', data_schedule)
#     for site in list(set(data_schedule['SITE_NAME'])):
#         new_data_sched = data_schedule[(data_schedule['SITE_NAME'] == site)]
#         duplicate_schedules = new_data_sched[new_data_sched.duplicated()]
#         if not duplicate_schedules.empty:
#             print("Duplicate schedules in ", site, ". Duplicate schedules are:")
#             print(duplicate_schedules)
#
#     newStopQuery = StopQuery()
#     stops_query, params_stops = newStopQuery.build_reg_stop_query(list_of_sites, type, date, [])
#
#     data_stops = pd.read_sql(stops_query, conn, params=p)
#     if data_stops.empty:
#         print("Oh no! Stops file for ", list_of_sites, " is empty!")
#     print_zb_files('Stops', list_of_sites, separate, 'SITE_NAME', data_stops)
#
#     run_vehicle_and_rot(sites_string, p, list_of_sites, separate, conn))


def run_diagnostic(list_of_sites, type, date):
    if type == 'Schedule Type':
        type = 'I'
    else:
        type = type_dict[type]
    if date == 'Date':
        date = ''
    conn = get_connection()
    diag_query = ScheduleQuery()
    diag_sched_query, diag_p = diag_query.build_reg_query(list_of_sites, type, date, '')
    data_schedule = pd.read_sql(diag_sched_query, conn, params=diag_p)

    diag_stop_query = StopQuery()
    diag_stop_query, diag_stop_p = diag_stop_query.build_reg_stop_query(list_of_sites, type, date, [])
    data_stop = pd.read_sql(diag_stop_query, conn, params=diag_stop_p)

    sites_string = ', '.join([str(':' + k) for k in diag_p.keys()])
    diag_rot_query = query_rot.replace('SITES_STRING_TO_REPLACE', sites_string)
    diag_data_rot = pd.read_sql(diag_rot_query, conn, params=diag_p)
    if diag_data_rot.empty:
        print("Oh no! ROT file for ", list_of_sites, " is empty!")

    print('Printing ROT file to HTML Check...')
    diag_data_rot.to_excel('HTML Check\\' + list_of_sites[0] + ' ROT ZB ' + time.strftime("%Y%m%d") + '.xlsx', index=False)

    print('Printing Times Route file to HTML Check...')
    # getting times route for ROT df
    data_times_for_rot = get_times_schedule_df(data_schedule, conn)
    data_times_for_rot.to_excel('HTML Check\\' + list_of_sites[0] + ' TIMES ROUTE ZB ' + time.strftime("%Y%m%d") + '.xlsx',
                                index=False)

    print('Printing schedules to HTML Check...')
    data_schedule.to_excel('HTML Check\\' + list_of_sites[0] + '_SCHEDULES ' + time.strftime("%Y%m%d") + '.xlsx',
                           index=False, encoding='Cp1252')
    print('Printing stops to HTML Check...')
    data_stop.to_excel('HTML Check\\' + list_of_sites[0] + '_STOPS ' + time.strftime("%Y%m%d") + '.xlsx', index=False,
                       encoding='Cp1252')
    # closing connection after printing all
    conn.close()
    print('Connection closed')
    db_files = []
    for file in os.listdir('HTML Check/'):
        if file.endswith('.xls') or file.endswith('.xlsx') or file.endswith('.XLSX'):
            db_files.append(str('HTML Check/' + file))
    print_diagnostic(db_files, '', 'from db', 'diagnostic.py')


if __name__ == "__main__":
    conn = get_connection()
    schedule_df = pd.read_excel("DataFiles/CLEVELAND SCHEDULES 20210922.xlsx")
    get_times_schedule_df(schedule_df, conn)