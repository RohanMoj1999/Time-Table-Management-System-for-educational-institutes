use pythonConnection

select * from examSchedule

if OBJECT_ID(N'classSchedule',N'U') is null begin create table classSchedule(Mon char(10) NOT NULL, Tues char(10) NOT NULL, TIME time(7) NOT NULL) end

select * from classSchedule

drop table classSchedule