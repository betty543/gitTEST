<% Option Explicit %>
<%
dim process_filename

process_filename = "./list05.asp"

dim Filename
Filename = "��������������Ȳ_" & Right(Replace(FormatDateTime(Date,2),"-",""),10) & "_data.xls"

Response.Buffer = True
Response.CacheControl = "public"
Response.ContentType = "application/vnd.ms-excel"
Response.AddHeader "Content-disposition","attachment;filename="&Filename

Server.execute(process_filename)
%>