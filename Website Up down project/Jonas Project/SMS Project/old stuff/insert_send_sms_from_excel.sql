Alter procedure insert_send_sms_from_excel
@sent_message nvarchar(4000),
@mobile_no nvarchar(1000),
@guid nvarchar(4000),
@status nvarchar(4000),
@reponse nvarchar(4000)
as
begin




Declare @institution_id int =3
Declare @school_id int =12
Declare @academic_year int =2019
Declare @sender_id  nvarchar(200)='CMSLKO'
Declare @notification nvarchar(200)='notification'



insert into notification_sent_sms_info 
values
(
@institution_id,

@school_id,
@academic_year,
@sender_id,
0,
@sent_message,
@mobile_no,
convert(varchar, getdate(), 0) ,
@notification,
@guid,
getdate(),
1,
null,
@status,
null,
1,
getdate(),
1,
getdate(),
0,1

)


end