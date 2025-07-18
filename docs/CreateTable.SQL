create table [dbo].[baseworktimetbl] (
  [id] int identity not null
  , [updatetime] timestamp not null
  , [personalcode] char(5)
  , [ymb] char(6)
  , [basemin] int
  , primary key (id)
);


create table [dbo].[controltbl] (
  [id] int not null
  , [updatetime] timestamp not null
  , [systemenable] char(1)
  , primary key (id)
);


create table [dbo].[deductiontbl] (
  [id] int identity not null
  , [updatetime] timestamp not null
  , [personalcode] char(5)
  , [ymb] char(6)
  , [amount01] int
  , [amount02] int
  , [amount03] int
  , [amount04] int
  , [amount05] int
  , [amount06] int
  , [amount07] int
  , [amount08ncr] char(10)
  , [amount08] int
  , [amount09ncr] char(10)
  , [amount09] int
  , [amount10ncr] char(10)
  , [amount10] int
  , primary key (id)
);


create table [dbo].[dutyrostertbl] (
  [id] int identity not null
  , [updatetime] timestamp not null
  , [personalcode] char(5)
  , [ymb] char(6)
  , [workdays] float(53)
  , [workholidays] float(53)
  , [absencedays] float(53)
  , [paidvacations] float(53)
  , [preservevacations] float(53)
  , [specialvacations] float(53)
  , [holidayshifts] float(53)
  , [realworkdays] float(53)
  , [shortdays] float(53)
  , [nightduty_a] int
  , [nightduty_b] int
  , [nightduty_c] int
  , [nightduty_d] int
  , [holidaypremium] float(53)
  , [dayduty] int
  , [shiftwork_kou] float(53)
  , [shiftwork_otsu] float(53)
  , [shiftwork_hei] float(53)
  , [summons] int
  , [summonslate] int
  , [yearend1230] float(53)
  , [yearend1231] float(53)
  , [workholidaytime] float(53)
  , [latepremium] float(53)
  , [overtime] float(53)
  , [holidayshifttime] float(53)
  , [holidayshiftovertime] float(53)
  , [holidayshiftlate] float(53)
  , [overtimelate] float(53)
  , [holidayshiftovertimelate] float(53)
  , [vacationnumber] float(53)
  , [holidaynumber] float(53)
  , [vacationtime] float(53)
  , [shiftwork_a] float(53)
  , [shiftwork_b] float(53)
  , [saturday_workmin] float(53)
  , [weekdays_workmin] float(53)
  , [workingmins] int
  , [currentworkmin] int
  , [legalholiday_extra_min] int
  , [weekovertime] float(53)
  , primary key (id)
);


create table [dbo].[holidaytbl] (
  [id] int identity not null
  , [updatetime] timestamp not null
  , [holidaydate] char(8)
  , [holidaytype] char(1)
  , [memo] char(100)
  , primary key (id)
);


create table [dbo].[iptbl] (
  [updatetime] timestamp not null
  , [ipnumber] char(15)
  , [personalcode] char(10)
  , [begindate] char(8)
  , [enddate] char(8)
);


create table [dbo].[orgtbl] (
  [id] int identity not null
  , [updatetime] timestamp not null
  , [personalcode] char(5)
  , [manageclass] char(1)
  , [orgcode] char(6)
  , primary key (id)
);


create table [dbo].[pctimetbl] (
  [updatetime] timestamp not null
  , [personalcode] char(10)
  , [ipnumber] char(15)
  , [pcdate] char(8)
  , [pctime] char(4)
  , [pcstatus] char(10)
);


create table [dbo].[remainvacationtbl] (
  [updatetime] timestamp not null
  , [personalcode] char(10)
  , [ymb] char(6)
  , [remainvacation] float(53)
);


create table [dbo].[stafftbl] (
  [id] int identity not null
  , [updatetime] timestamp not null
  , [personalcode] char(5)
  , [staffname] char(30)
  , [orgcode] char(6)
  , [gradecode] char(3)
  , [is_operator] char(1)
  , [is_input] char(1)
  , [is_charge] char(1)
  , [is_superior] char(1)
  , [is_enable] char(1)
  , [password] char(42)
  , [processed_ymb] char(6)
  , [holidaytype] char(1)
  , [is_deduction] char(1)
  , [opentime] char(4)
  , [closetime] char(4)
  , [is_unionexecutive] char(1)
  , [grantdate] char(4)
  , [workshift] char(1) default '0'
  , [old_workshift] char(1)
  , [old_workshift_last_ymb] char(6) default '000000'
  , [base_am_workmin] int
  , [base_pm_workmin] int
  , primary key (id)
);


create table [dbo].[timetbl] (
  [id] int identity not null
  , [updatetime] timestamp not null
  , [personalcode] char(5)
  , [workingdate] char(8)
  , [comedate] char(8)
  , [cometime] char(4)
  , [outdate] char(8)
  , [outtime] char(4)
  , [returndate] char(8)
  , [returntime] char(4)
  , [leavedate] char(8)
  , [leavetime] char(4)
  , primary key (id)
);


create table [dbo].[typetbl] (
  [id] int identity not null
  , [codetype] char(20)
  , [uppercode] char(20)
  , [code] char(20)
  , [codetext] char(40)
  , [abbrcodetext] char(20)
  , [dispseq] int
  , primary key (id)
);


create table [dbo].[worktbl] (
  [id] int identity not null
  , [updatetime] timestamp not null
  , [personalcode] char(5)
  , [workingdate] char(8)
  , [morningwork] char(1)
  , [afternoonwork] char(1)
  , [morningholiday] char(1)
  , [afternoonholiday] char(1)
  , [summons] char(1)
  , [overtime_begin] char(4)
  , [overtime_end] char(4)
  , [rest_begin] char(4)
  , [rest_end] char(4)
  , [overtime] char(4)
  , [overtimelate] char(4)
  , [holidayshift] char(4)
  , [holidayshiftovertime] char(4)
  , [holidayshiftlate] char(4)
  , [holidayshiftovertimelate] char(4)
  , [requesttime] char(4)
  , [requesttime_begin] char(4)
  , [requesttime_end] char(4)
  , [latetime] char(4)
  , [latetime_begin] char(4)
  , [latetime_end] char(4)
  , [is_approval] char(1)
  , [nightduty] char(1)
  , [dayduty] char(1)
  , [operator] char(1)
  , [vacationtime] char(4)
  , [vacationtime_begin] char(4)
  , [vacationtime_end] char(4)
  , [memo] char(100)
  , [memo2] char(2)
  , [is_error] char(1)
  , [work_begin] char(4)
  , [work_end] char(4)
  , [break_begin1] char(4)
  , [break_end1] char(4)
  , [break_begin2] char(4)
  , [break_end2] char(4)
  , [workmin] int default 0
  , [weekovertime] char(4)
  , primary key (id)
);

