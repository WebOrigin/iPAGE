<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<% Server.ScriptTimeOut=5000 %>
<!--#include FILE="upload.inc"--> 

<% 


dim upload,file,formName,formPath 
set upload=new upload_5xSoft





cufname = upload.form("First Name")
culname = upload.form("Last Name")
cumail = upload.form("E-mail")
cucom = upload.form("Company")
cuphone = upload.form("Phone")
cuclass = upload.form("Classification")
cusletter = upload.form("Newsletter")
cusmsg = upload.form("Message")
cusatt = upload.form("Supporting Material")






formPath=upload.form("filepath")
if right(formPath,1)<>"/" then formPath=formPath&"/" 
for each formName in upload.file
set file=upload.file(formName)
if file.FileSize>0 then
file_rst=Server.mappath("mail_upload\"&file.FileName)
file.SaveAs file_rst
end if 
set file=nothing 
next 
set upload=nothing 

myemail = "contact@ipagenz.co.nz"
myname = "iPAGE New Zealand"
smtpserver ="mail.ipagenz.co.nz"
smtpuser = "webcustomer@ipagenz.co.nz"
smtppwd = "345gtrg45ersd3g544"
smtpemail = "webcustomer@ipagenz.co.nz"

set jmail = server.CreateObject ("jmail.message")

jmail.From = "webcustomer@ipagenz.co.nz"
jmail.FromName = cufname & " " & culname
jmail.ReplyTo = cumail
jmail.Subject = "Web Customer: " & cufname & " " & culname

jmail.Body ="Name: " & cufname & " " & culname & vbcrlf & vbcrlf & "E-mail: " & cumail & vbcrlf & vbcrlf & "Company: " & cucom & vbcrlf & vbcrlf & "Phone: " & cuphone & vbcrlf & vbcrlf & "Classification: " & cuclass & vbcrlf & vbcrlf & "Newsletter: " & cusletter & vbcrlf & vbcrlf & "Message:" & vbcrlf & cusmsg

jmail.AddAttachment file_rst
jmail.AddRecipient myemail,myname
jmail.MailServerUserName = smtpuser
jmail.MailServerPassWord = smtppwd

isgo=jmail.Send(smtpserver)
if isgo then
Response.Redirect("http://www.ipagenz.co.nz/respond_success.html")
else
Response.Redirect("http://www.ipagenz.co.nz/respond_fail.html")
end if

jmail.Close
set jmail=nothing
%>