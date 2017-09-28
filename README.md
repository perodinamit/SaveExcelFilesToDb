# SaveExcelFilesToDb
Reads excel files from folder and saves them into database


1. first create simple database table like:

CREATE TABLE [dbo].[tblCustomer](
    [id] int NULL,
    [name] varchar(100) NULL,
    [dob] date NULL
) 

select * from [dbo].[tblCustomer]
