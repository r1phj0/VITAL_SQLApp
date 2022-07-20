import datetime
import pandas as pd


query_vehicles = """with vt_des as (select * from vitalprod.code_value \nwhere code_type_name = 'Vehicle Type'), \nvs_des as (select * from vitalprod.code_value \nwhere code_type_name = 'Vehicle Status') \nselect  f.area_name, f.fac_id, f.fac_name,  v.veh_id,
    vs_des.decode_desc vehicle_status,
    v.status_dtm, v.yr_nbr, vt_des.decode_desc vehicle_type, curr_month_begin_mileage_nbr
    from VITALPROD.FACILITY_T f
    left join VITALPROD.VEHICLE_T v on v.fac_id  = f.fac_id
    left join vt_des on vt_des.code = v.veh_type_cd \nleft join vs_des on vs_des.code = v.vsc_status_cd \nwhere f.fac_name in (SITES_STRING_TO_REPLACE) \nand f.PVS_SITE_IND = 'Y' \norder by vehicle_status"""
query_rot = """with orig as (select fac_id,fac_type_cd, fac_name, LINE1_ADDR, line2_addr, City_name, state_id, nass_cd from vitalprod.facility_t),
    dest as (select fac_id, fac_type_cd, fac_name,LINE1_ADDR, line2_addr, City_name, state_id, nass_cd from vitalprod.facility_t),
    pvs_site as (select fac_id, fac_name from vitalprod.facility_t) 

    select  orig_fac_id, orig.fac_type_cd orig_type, dest_fac_id,dest.fac_type_cd dest_type, pvs_site_id,orig.fac_name orig_name, dest.fac_name dest_name, 
    orig.LINE1_ADDR orig_addr1, orig.line2_addr orig_addr2, orig.City_name orig_city, orig.state_id orig_state,
    dest.LINE1_ADDR dest_addr1, dest.line2_addr dest_addr2, dest.City_name dest_city, dest.state_id dest_state,
    orig.nass_cd orig_nass, dest.nass_cd dest_nass,
    mileage_nbr, PVS_SITE.FAC_NAME PVS_NAME,DIR_DESC
    from vitalprod.travel_t tt
    left join orig on orig.fac_id = orig_fac_id
    left join dest on dest.fac_id = dest_fac_id
    left join pvs_site on pvs_site.fac_id = pvs_site_id

    where pvs_site.fac_name in (SITES_STRING_TO_REPLACE)
    order by orig.fac_name, dest.fac_name"""
times_sched_for_rot = """with so as (select route_id, transp_trip_id, effective_dt, dispatch_leg_nbr, scheduled_time, 
ref_leg_fac_cd
from vitalprod.TIMES_SCHEDULE_T
where trip_direction_ind = 'O')


select si.route_id, si.transp_trip_id,  si.effective_dt, 
si.dispatch_leg_nbr,si.ref_leg_fac_cd orig_nass, so.ref_leg_fac_cd dest_nass

from vitalprod.TIMES_SCHEDULE_T si
left join so on si.route_id = so.route_id and si.transp_trip_id = so.transp_trip_id 
and si.effective_dt = so.effective_dt and si.dispatch_leg_nbr = so.dispatch_leg_nbr
WHERE si.ROUTE_ID LIKE 'ROUTE_ID_REPLACE' -- input
AND si.TRANSP_TRIP_ID LIKE '%TRANSP_TRIP_ID_Spec%'--input
AND si.EFFECTIVE_DT LIKE 'EFF_DATE%'--input
and si.trip_direction_ind = 'I' -- don't change
order by si.route_id, si.transp_trip_id
"""


def get_times_sched_for_rot(replace_list, conn):
    temp = times_sched_for_rot.replace('ROUTE_ID_REPLACE', replace_list[0])
    temp1 = temp.replace('TRANSP_TRIP_ID_Spec', replace_list[1])
    temp2 = temp1.replace('EFF_DATE', replace_list[2])
    df = pd.read_sql(temp2, conn)
    df['SITE_NAME'] = replace_list[3]
    return df


class ScheduleQuery:
    def __init__(self):
        self.sites_string_beginning_part = """with cd_des as (select * from vitalprod.code_value \nwhere code_type_name = 'Schedule Status') \nselect  
                         f.area_name, s.pvs_site_id,f.fac_name site_name,  cd_des.decode_desc, s.sch_sched_nbr,  s.tour_nbr, s.sched_type_id, s.run_nbr, s.frq_cd, 
                        s.tractor_ind, s.veh_1_id, s.veh_2_id, s.veh_3_id,
                        s.mileage_nbr,
                        to_char(s.begin_time, 'HH:MI AM') start_time, to_char(s.end_time,'HH:MI AM') end_time ,
                        substr(to_char((s.end_time - s.begin_time + interval '24' hour), 'DD HH:MI'),12,5) sch_duration,
                        s.tot_stop_cnt,    
                        s.sch_effect_dtm, s.end_dt, 
                        s.SCH_ROUTE_ID, s.trip_cd,
                        s.route_prev_ind, s.sched_no_prev_ind,
                        s.end_times_dt, s.inv_times_ind, s.time_trip_id,s.unschedule_trip_ind,
                        s.show_start_time_ind,
                        s.comment_text, s.ver_nbr
                        
                        from 
                        VITALPROD.SCHEDULE_T s
                        left join vitalprod.facility_t f on f.fac_id = s.PVS_SITE_ID
                        left join cd_des on cd_des.code = s.ssc_status_cd
                        where
                        f.fac_name in (SITES_STRING_TO_REPLACE)
                        AND F.PVS_SITE_IND = 'Y'
                        """
        self.schedule_numbers = """\nAND s.SCH_SCHED_NBR in (SCHEDULES_STRING_TO_REPLACE)"""
        self.type_string = """\nAND cd_des.attr_1_text in ('TYPE_LETTER_TO_REPLACE')"""
        self.date_string = """\nand (to_char(s.sch_effect_dtm, 'DD-MM-YYYY') = to_char(to_date('DATE_STRING_TO_REPLACE', 
        'DD-MM-YYYY'), 'DD-MM-YYYY'))"""
        self.end_string = """\norder by  F.FAC_NAME, S.sch_sched_nbr, S.RUN_NBR, cd_des.attr_1_text"""
        self.schedule_count_query = """"""
        self.final_query = ''

    def build_reg_query(self, sites, type, date, schedule_nums):
        p = {str('site' + str(sites.index(i))): i for i in sites}
        sites_string = ', '.join([str(':' + k) for k in p.keys()])
        beg_and_sites_string = self.sites_string_beginning_part.replace('SITES_STRING_TO_REPLACE', sites_string)
        type_string = self.type_string.replace('TYPE_LETTER_TO_REPLACE', type)

        if date != '':
            date_string = self.date_string.replace('DATE_STRING_TO_REPLACE', date)
        else:
            date_string = ''
        if len(schedule_nums) != 0:
            schedules_list = ', '.join([str("'" + k + "'") for k in schedule_nums])
            schedule_string = self.schedule_numbers.replace('SCHEDULES_STRING_TO_REPLACE', schedules_list)
        else:
            schedule_string = ''

        self.final_query = beg_and_sites_string + type_string + date_string + schedule_string + self.end_string

        return self.final_query, p

    def build_zb_spec_query(self, site, type, date, schedule_nums):
        zb_spec_params = {str('sched' + str(schedule_nums.index(i))): str(i) for i in schedule_nums}
        schedules_list = ', '.join([str(':' + k) for k in zb_spec_params.keys()])
        schedule_string = self.schedule_numbers.replace('SCHEDULES_STRING_TO_REPLACE', schedules_list)
        beg_and_sites_string = self.sites_string_beginning_part.replace('SITES_STRING_TO_REPLACE', str("'" + site + "'"))
        type_string = self.type_string.replace('TYPE_LETTER_TO_REPLACE', type)

        if date != '':
            date_string = self.date_string.replace('DATE_STRING_TO_REPLACE', date)
        else:
            date_string = ''

        self.final_query = beg_and_sites_string + schedule_string + type_string + date_string + self.end_string

        return self.final_query, zb_spec_params

    # used when there are schedules missing - pulls all schedule information in VITAL for one site
    def build_missing_sched_query(self, site):
        beg_and_sites_string = self.sites_string_beginning_part.replace('SITES_STRING_TO_REPLACE', str("'" + site + "'"))

        self.final_query = beg_and_sites_string + self.end_string

        return self.final_query

    def build_count_query(self, times_route, times_trip, freq_code):
        times_trip = str(times_trip + '%')


class StopQuery:
    def __init__(self):
        self.beg_w_site_string = """
            with schedule as (select f.area_name, s.pvs_site_id,
            f.fac_name, s.sch_route_id, s.time_trip_id, s.ssc_status_cd, cd_des.decode_desc,
            s.sch_sched_nbr, s.SCH_EFFECT_DTM, s.LAST_UPDT_DTM
            from VITALPROD.SCHEDULE_T s
            left join vitalprod.facility_t f on f.fac_id = s.PVS_SITE_ID
            left join vitalprod.code_value cd_des on cd_des.code = s.ssc_status_cd
            where
            cd_des.code_type_name = 'Schedule Status'
            and f.fac_name in (SITES_STRING_TO_REPLACE)"""
        self.schedule_num_str = """\nAND s.SCH_SCHED_NBR in (SCHEDULES_STRING_TO_REPLACE)"""
        self.type_string = """\nAND cd_des.attr_1_text in ('TYPE_LETTER_TO_REPLACE')"""
        self.date_string = """\nand (to_char(s.sch_effect_dtm, 'DD-MM-YYYY') = to_char(to_date('DATE_STRING_TO_REPLACE', 
                'DD-MM-YYYY'), 'DD-MM-YYYY'))"""
        self.stop_end_string = """
            and f.pvs_site_ind = 'Y')
            SELECT schedule.area_name, schedule.pvs_site_id, schedule.fac_name site_name,
            f.fac_name stop_name, f.nass_cd, sp.fac_id stop_id,
            schedule.decode_desc, sp.sch_sched_nbr, schedule.time_trip_id, sp.frq_cd,
            to_char(sp.arr_time,'HH:MI AM') arr_time,
            to_char(sp.dep_time, 'HH:MI AM') dep_time,
            sp.stop_nbr, sp.oper_instr_text,
            f.line1_addr, f.line2_addr, f.city_name, f.state_ID
            from vitalprod.SERVICE_PT_T sp
            join schedule on schedule.sch_sched_nbr = sp.sch_sched_nbr
            left join vitalprod.facility_t f on f.FAC_ID = sp.fac_id
            where
            sp.sch_route_id in (schedule.sch_route_id )
            and sp.sch_sched_nbr in (schedule.sch_sched_nbr)
            and sp.SCH_EFFECT_DTM in (schedule.SCH_EFFECT_DTM )
            order by schedule.FAC_NAME, schedule.decode_desc, sp.sch_sched_nbr, sp.stop_nbr"""
        self.final_query = ''

    def build_reg_stop_query(self, sites, type, date, schedule_nums):
        stop_params = {str('site' + str(sites.index(i))): i for i in sites}
        sites_string = ', '.join([str(':' + k) for k in stop_params.keys()])
        beg_and_sites_string = self.beg_w_site_string.replace('SITES_STRING_TO_REPLACE', sites_string)
        type_string = self.type_string.replace('TYPE_LETTER_TO_REPLACE', type)

        if date != '':
            date_string = self.date_string.replace('DATE_STRING_TO_REPLACE', date)
        else:
            date_string = ''
        if len(schedule_nums) != 0:
            schedules_list = ', '.join([str("'" + k + "'") for k in schedule_nums])
            schedule_string = self.schedule_num_str.replace('SCHEDULES_STRING_TO_REPLACE', schedules_list)
        else:
            schedule_string = ''
        self.final_query = beg_and_sites_string + type_string + date_string + schedule_string + self.stop_end_string

        return self.final_query, stop_params

    def build_zb_spec_stop_query(self, site, schedule_num, type, date):
        if type== 'DS':
            type = 'D'
        beg_and_sites_string = self.beg_w_site_string.replace('SITES_STRING_TO_REPLACE', str("'" + site + "'"))
        type_string = self.type_string.replace('TYPE_LETTER_TO_REPLACE', type)
        if date != '':
            date_string = self.date_string.replace('DATE_STRING_TO_REPLACE', date)
        else:
            date_string = ''
        schedule_string = self.schedule_num_str.replace('SCHEDULES_STRING_TO_REPLACE', str("'" + schedule_num + "'"))
        self.final_query = beg_and_sites_string + type_string + date_string + schedule_string + self.stop_end_string

        return self.final_query


class ScheduleValidation:
    def __init__(self, zb_row, dates):
        # info from ZB workbook
        self.schedule_num = zb_row['Schedule #']
        self.zb_type = zb_row['Type (D / DS / I / F)']
        if self.zb_type == 'D':
            self.date = dates[0]
        elif self.zb_type == 'DS':
            self.date = dates[1]
        elif self.zb_type == 'F':
            self.date = dates[2]
        else:
            self.date = ''
        self.zb_freq = zb_row['FREQ']
        self.zb_workhr = zb_row['Paid Dly Hrs']
        self.zb_annual_mileage = zb_row['Annual Mileage']
        self.zb_annual_hours = zb_row['Annual Work Hours']
        # info from SQL data pull
        self.vital_freq = ''
        self.vital_workhr = ''
        self.vital_annual_mileage = 0
        self.vital_annual_hours = ''
        self.vital_sched_duration = 0
        self.stops_df = None
        self.total_duration = 0
        self.paid_time = ''
        self.lunch_time = ''
        self.final_db = None


def stop_time_mismatch(stops_df):
    # stops_df = pd.read_excel(r'C:\Users\R1PHJ0\PycharmProjects\VITAL_SQLDataApp\DataFiles\QUEENS STOPS ZB 20210823.xlsx')

    stops_df['ARR_TIME'] = stops_df['ARR_TIME'].apply(lambda x: x.strftime('%I:%M %p'))
    stops_df['DEP_TIME'] = stops_df['DEP_TIME'].apply(lambda x: x.strftime('%I:%M %p'))
    problem_stops = pd.DataFrame(columns=['Schedule #', 'Stop #', 'Start Time', 'End Time', 'Indicator'])

    list_of_schedules = set(list(stops_df['SCH_SCHED_NBR']))

    for schedule in list_of_schedules:
        pass_midnight = False
        occurance = 0
        previous_stop = ''
        previous_start = ''
        previous_end = ''
        previous_end_adj = ''
        previous_old_end_time = ''
        previous_end_ap = ''
        day_start = 1
        day_end = 1
        temp_df = stops_df[stops_df['SCH_SCHED_NBR'] == schedule]
        temp_df.reset_index(drop=True, inplace=True)
        for stop in range(temp_df.shape[0]):
            stop_number = temp_df.loc[stop, 'STOP_NBR']
            old_start_time = pd.to_datetime(temp_df.loc[stop, 'ARR_TIME'])
            start_ap = str(temp_df.loc[stop, 'ARR_TIME'])[-2:]
            old_end_time = pd.to_datetime(temp_df.loc[stop, 'DEP_TIME'])
            end_ap = str(temp_df.loc[stop, 'DEP_TIME'])[-2:]

            if pass_midnight:
                if start_ap == 'PM' or end_ap == 'PM':
                    day_start = 1
                    day_end = 1
                    pass_midnight = False
                else:
                    day_start = 2
                    day_end = 2

            if stop_number != 1:
                # print((old_start_time-previous_old_end_time).total_seconds()/60)
                if (old_end_time - old_start_time).total_seconds() / 60 < 1 and not pass_midnight:
                    day_start = 1
                    day_end = 2
                    pass_midnight = True
                if (old_start_time - previous_old_end_time).total_seconds() / 60 < 0:
                    day_start = 2
                    day_end = 2
                    pass_midnight = True
            try:
                start_time = datetime.datetime.strptime(str(temp_df.loc[stop, 'ARR_TIME']), '%I:%M %p').replace(
                    year=2022,
                    month=1,
                    day=day_start)
                end_time = datetime.datetime.strptime(str(temp_df.loc[stop, 'DEP_TIME']), '%I:%M %p').replace(year=2022,
                                                                                                              month=1,
                                                                                                              day=day_end)
            except:
                start_time = datetime.datetime.strptime(str(temp_df.loc[stop, 'ARR_TIME'][:15] + ' ' +
                                                            temp_df.loc[stop, 'ARR_TIME'][-2:]).replace('.', ':'),
                                                        '%d-%b-%y %I:%M %p').replace(year=2022, month=1, day=day_start)
                end_time = datetime.datetime.strptime(str(temp_df.loc[stop, 'DEP_TIME'][:15] + ' ' +
                                                          temp_df.loc[stop, 'DEP_TIME'][-2:]).replace('.', ':'),
                                                      '%d-%b-%y %I:%M %p').replace(year=2022, month=1, day=day_end)
            # if (end_time - start_time).total_seconds() / 60 < 1:
            #     adj_end_time  = end_time + datetime.timedelta(hours = 12)
            #     pass_midnight = True
            # else:
            #     adj_end_time  = end_time
            adj_end_time = end_time
            if stop_number == 1:
                previous_stop = stop_number
                previous_start = start_time
                previous_end = end_time
                previous_end_adj = adj_end_time
                previous_old_end_time = old_end_time
            else:
                # print(f"stop # : {stop_number}, {(start_time-previous_end_adj).total_seconds()/60}")
                if (start_time - previous_end_adj).total_seconds() / 60 > 480 or (
                        start_time - previous_end_adj).total_seconds() / 60 < 0:
                    occurance += 1
                    row = problem_stops.shape[0] + 1
                    problem_stops.loc[row, 'Schedule #'] = schedule
                    problem_stops.loc[row, 'Stop #'] = previous_stop
                    problem_stops.loc[row, 'Start Time'] = previous_start.strftime('%I:%M %p')
                    problem_stops.loc[row, 'End Time'] = previous_end.strftime('%I:%M %p')
                    problem_stops.loc[row, 'Indicator'] = occurance

                    next_row = row + 1
                    problem_stops.loc[next_row, 'Schedule #'] = schedule
                    problem_stops.loc[next_row, 'Stop #'] = stop_number
                    problem_stops.loc[next_row, 'Start Time'] = start_time.strftime('%I:%M %p')
                    problem_stops.loc[next_row, 'End Time'] = end_time.strftime('%I:%M %p')
                    problem_stops.loc[next_row, 'Indicator'] = occurance

                previous_stop = stop_number
                previous_start = start_time
                previous_end = end_time
                previous_end_adj = adj_end_time
                previous_old_end_time = old_end_time

    return problem_stops
