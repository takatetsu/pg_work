-- PostgreSQL用テーブル作成スクリプト
-- SQL Serverから変換

-- pgcrypto拡張機能を有効化（ハッシュ関数用）
CREATE EXTENSION IF NOT EXISTS pgcrypto;

-- 基準労働時間テーブル
CREATE TABLE baseworktimetbl (
    id SERIAL PRIMARY KEY,
    updatetime TIMESTAMP WITH TIME ZONE NOT NULL DEFAULT CURRENT_TIMESTAMP,
    personalcode CHAR(5),
    ymb CHAR(6),
    basemin INTEGER
);
CREATE INDEX idx_baseworktimetbl_personalcode ON baseworktimetbl(personalcode);
CREATE INDEX idx_baseworktimetbl_ymb ON baseworktimetbl(ymb);
COMMENT ON TABLE baseworktimetbl IS '基準労働時間テーブル';

-- コントロールテーブル
CREATE TABLE controltbl (
    id INTEGER NOT NULL PRIMARY KEY,
    updatetime TIMESTAMP WITH TIME ZONE NOT NULL DEFAULT CURRENT_TIMESTAMP,
    systemenable CHAR(1)
);
COMMENT ON TABLE controltbl IS 'システム制御テーブル';

-- 支店控除テーブル
CREATE TABLE deductiontbl (
    id SERIAL PRIMARY KEY,
    updatetime TIMESTAMP WITH TIME ZONE NOT NULL DEFAULT CURRENT_TIMESTAMP,
    personalcode CHAR(5),
    ymb CHAR(6),
    amount01 INTEGER,
    amount02 INTEGER,
    amount03 INTEGER,
    amount04 INTEGER,
    amount05 INTEGER,
    amount06 INTEGER,
    amount07 INTEGER,
    amount08ncr CHAR(10),
    amount08 INTEGER,
    amount09ncr CHAR(10),
    amount09 INTEGER,
    amount10ncr CHAR(10),
    amount10 INTEGER
);
CREATE INDEX idx_deductiontbl_personalcode ON deductiontbl(personalcode);
CREATE INDEX idx_deductiontbl_ymb ON deductiontbl(ymb);
COMMENT ON TABLE deductiontbl IS '支店控除テーブル';

-- 勤務表テーブル
CREATE TABLE dutyrostertbl (
    id SERIAL PRIMARY KEY,
    updatetime TIMESTAMP WITH TIME ZONE NOT NULL DEFAULT CURRENT_TIMESTAMP,
    personalcode CHAR(5),
    ymb CHAR(6),
    workdays DOUBLE PRECISION,
    workholidays DOUBLE PRECISION,
    absencedays DOUBLE PRECISION,
    paidvacations DOUBLE PRECISION,
    preservevacations DOUBLE PRECISION,
    specialvacations DOUBLE PRECISION,
    holidayshifts DOUBLE PRECISION,
    realworkdays DOUBLE PRECISION,
    shortdays DOUBLE PRECISION,
    nightduty_a INTEGER,
    nightduty_b INTEGER,
    nightduty_c INTEGER,
    nightduty_d INTEGER,
    holidaypremium DOUBLE PRECISION,
    dayduty INTEGER,
    shiftwork_kou DOUBLE PRECISION,
    shiftwork_otsu DOUBLE PRECISION,
    shiftwork_hei DOUBLE PRECISION,
    summons INTEGER,
    summonslate INTEGER,
    yearend1230 DOUBLE PRECISION,
    yearend1231 DOUBLE PRECISION,
    workholidaytime DOUBLE PRECISION,
    latepremium DOUBLE PRECISION,
    overtime DOUBLE PRECISION,
    holidayshifttime DOUBLE PRECISION,
    holidayshiftovertime DOUBLE PRECISION,
    holidayshiftlate DOUBLE PRECISION,
    overtimelate DOUBLE PRECISION,
    holidayshiftovertimelate DOUBLE PRECISION,
    vacationnumber DOUBLE PRECISION,
    holidaynumber DOUBLE PRECISION,
    vacationtime DOUBLE PRECISION,
    shiftwork_a DOUBLE PRECISION,
    shiftwork_b DOUBLE PRECISION,
    saturday_workmin DOUBLE PRECISION,
    weekdays_workmin DOUBLE PRECISION,
    workingmins INTEGER,
    currentworkmin INTEGER,
    legalholiday_extra_min INTEGER,
    weekovertime DOUBLE PRECISION
);
CREATE INDEX idx_dutyrostertbl_personalcode ON dutyrostertbl(personalcode);
CREATE INDEX idx_dutyrostertbl_ymb ON dutyrostertbl(ymb);
COMMENT ON TABLE dutyrostertbl IS '勤務表テーブル（月次集計）';

-- 公休日テーブル
CREATE TABLE holidaytbl (
    id SERIAL PRIMARY KEY,
    updatetime TIMESTAMP WITH TIME ZONE NOT NULL DEFAULT CURRENT_TIMESTAMP,
    holidaydate CHAR(8),
    holidaytype CHAR(1),
    memo CHAR(100)
);
CREATE INDEX idx_holidaytbl_holidaydate ON holidaytbl(holidaydate);
COMMENT ON TABLE holidaytbl IS '公休日テーブル';

-- IPテーブル
CREATE TABLE iptbl (
    updatetime TIMESTAMP WITH TIME ZONE NOT NULL DEFAULT CURRENT_TIMESTAMP,
    ipnumber CHAR(15),
    personalcode CHAR(10),
    begindate CHAR(8),
    enddate CHAR(8)
);
COMMENT ON TABLE iptbl IS 'IP管理テーブル';

-- 組織テーブル
CREATE TABLE orgtbl (
    id SERIAL PRIMARY KEY,
    updatetime TIMESTAMP WITH TIME ZONE NOT NULL DEFAULT CURRENT_TIMESTAMP,
    personalcode CHAR(5),
    manageclass CHAR(1),
    orgcode CHAR(6)
);
CREATE INDEX idx_orgtbl_personalcode ON orgtbl(personalcode);
CREATE INDEX idx_orgtbl_orgcode ON orgtbl(orgcode);
COMMENT ON TABLE orgtbl IS '組織権限テーブル';

-- 組織名テーブル
CREATE TABLE orgnametbl (
    orgcode CHAR(6) NOT NULL PRIMARY KEY,
    orgname CHAR(100)
);
COMMENT ON TABLE orgnametbl IS '組織名マスタテーブル';

-- 電源時刻テーブル
CREATE TABLE pctimetbl (
    updatetime TIMESTAMP WITH TIME ZONE NOT NULL DEFAULT CURRENT_TIMESTAMP,
    personalcode CHAR(10),
    ipnumber CHAR(15),
    pcdate CHAR(8),
    pctime CHAR(4),
    pcstatus CHAR(10)
);
COMMENT ON TABLE pctimetbl IS 'PC電源時刻テーブル';

-- 電源時刻MERGE用テーブル
CREATE TABLE pctimemergetbl (
    personalcode CHAR(10),
    ipnumber CHAR(15),
    pcdate CHAR(8),
    pctime CHAR(4),
    pcstatus CHAR(10)
);
COMMENT ON TABLE pctimemergetbl IS 'PC電源時刻統合用テーブル';

-- 保存休暇残日数テーブル
CREATE TABLE remainvacationtbl (
    updatetime TIMESTAMP WITH TIME ZONE NOT NULL DEFAULT CURRENT_TIMESTAMP,
    personalcode CHAR(10),
    ymb CHAR(6),
    remainvacation DOUBLE PRECISION
);
COMMENT ON TABLE remainvacationtbl IS '保存休暇残日数テーブル';

-- 社員テーブル
CREATE TABLE stafftbl (
    id SERIAL PRIMARY KEY,
    updatetime TIMESTAMP WITH TIME ZONE NOT NULL DEFAULT CURRENT_TIMESTAMP,
    personalcode CHAR(5),
    staffname CHAR(30),
    orgcode CHAR(6),
    gradecode CHAR(3),
    is_operator CHAR(1),
    is_input CHAR(1),
    is_charge CHAR(1),
    is_superior CHAR(1),
    is_enable CHAR(1),
    password CHAR(42),
    processed_ymb CHAR(6),
    holidaytype CHAR(1),
    is_deduction CHAR(1),
    opentime CHAR(4),
    closetime CHAR(4),
    is_unionexecutive CHAR(1),
    grantdate CHAR(4),
    workshift CHAR(1) DEFAULT '0',
    old_workshift CHAR(1),
    old_workshift_last_ymb CHAR(6) DEFAULT '000000',
    base_am_workmin INTEGER,
    base_pm_workmin INTEGER
);
CREATE INDEX idx_stafftbl_personalcode ON stafftbl(personalcode);
CREATE INDEX idx_stafftbl_orgcode ON stafftbl(orgcode);
COMMENT ON TABLE stafftbl IS '社員マスタテーブル';

-- タイムテーブル
CREATE TABLE timetbl (
    id SERIAL PRIMARY KEY,
    updatetime TIMESTAMP WITH TIME ZONE NOT NULL DEFAULT CURRENT_TIMESTAMP,
    personalcode CHAR(5),
    workingdate CHAR(8),
    comedate CHAR(8),
    cometime CHAR(4),
    outdate CHAR(8),
    outtime CHAR(4),
    returndate CHAR(8),
    returntime CHAR(4),
    leavedate CHAR(8),
    leavetime CHAR(4)
);
CREATE INDEX idx_timetbl_personalcode ON timetbl(personalcode);
CREATE INDEX idx_timetbl_workingdate ON timetbl(workingdate);
COMMENT ON TABLE timetbl IS 'タイムレコードテーブル（出退勤時刻）';

-- タイムカードテーブル
CREATE TABLE timecardtbl (
    updatetime TIMESTAMP WITH TIME ZONE NOT NULL DEFAULT CURRENT_TIMESTAMP,
    datasection CHAR(2),
    punchdate CHAR(8),
    punchtime CHAR(4),
    dutyclass CHAR(2),
    attendanceclass CHAR(2),
    personalcode CHAR(10),
    exceptioncode CHAR(2),
    terminalcode CHAR(2)
);
COMMENT ON TABLE timecardtbl IS 'タイムカードデータテーブル';

-- タイプテーブル
CREATE TABLE typetbl (
    id SERIAL PRIMARY KEY,
    codetype CHAR(20),
    uppercode CHAR(20),
    code CHAR(20),
    codetext CHAR(40),
    abbrcodetext CHAR(20),
    dispseq INTEGER
);
COMMENT ON TABLE typetbl IS 'コードマスタテーブル';

-- 勤怠テーブル
CREATE TABLE worktbl (
    id SERIAL PRIMARY KEY,
    updatetime TIMESTAMP WITH TIME ZONE NOT NULL DEFAULT CURRENT_TIMESTAMP,
    personalcode CHAR(5),
    workingdate CHAR(8),
    morningwork CHAR(1),
    afternoonwork CHAR(1),
    morningholiday CHAR(1),
    afternoonholiday CHAR(1),
    summons CHAR(1),
    overtime_begin CHAR(4),
    overtime_end CHAR(4),
    rest_begin CHAR(4),
    rest_end CHAR(4),
    overtime CHAR(4),
    overtimelate CHAR(4),
    holidayshift CHAR(4),
    holidayshiftovertime CHAR(4),
    holidayshiftlate CHAR(4),
    holidayshiftovertimelate CHAR(4),
    requesttime CHAR(4),
    requesttime_begin CHAR(4),
    requesttime_end CHAR(4),
    latetime CHAR(4),
    latetime_begin CHAR(4),
    latetime_end CHAR(4),
    is_approval CHAR(1),
    nightduty CHAR(1),
    dayduty CHAR(1),
    operator CHAR(1),
    vacationtime CHAR(4),
    vacationtime_begin CHAR(4),
    vacationtime_end CHAR(4),
    memo CHAR(100),
    memo2 CHAR(2),
    is_error CHAR(1),
    work_begin CHAR(4),
    work_end CHAR(4),
    break_begin1 CHAR(4),
    break_end1 CHAR(4),
    break_begin2 CHAR(4),
    break_end2 CHAR(4),
    workmin INTEGER DEFAULT 0,
    weekovertime CHAR(4)
);
CREATE INDEX idx_worktbl_personalcode ON worktbl(personalcode);
CREATE INDEX idx_worktbl_workingdate ON worktbl(workingdate);
CREATE INDEX idx_worktbl_personalcode_date ON worktbl(personalcode, workingdate);
COMMENT ON TABLE worktbl IS '勤怠データテーブル（日次）';
