Alter procedure update_sms_status_from_excel_data

@guid nvarchar(4000),
@status nvarchar(4000),
@response nvarchar(4000)
as
begin

update notification_sent_sms_info set status_message=@status,response=@response,edited_date=getdate() where gu_id=@guid

end