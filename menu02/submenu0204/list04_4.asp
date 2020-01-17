<%
'<!--불만족현황 소속 -->
set oCmd5.ActiveConnection = db
oCmd5.CommandText = "armyinformix.dbo.StoredProcedures_4"
oCmd5.CommandType = adCmdStoredProc

iAction = "1"

set prm = oCmd5.CreateParameter("@iAction",adChar,adParamInput,1,iAction)
oCmd5.Parameters.Append prm
set prm = oCmd5.CreateParameter("@sDate",adChar,adParamInput,10,FromDate)
oCmd5.Parameters.Append prm
set prm = oCmd5.CreateParameter("@eDate",adChar,adParamInput,10,ToDate)
oCmd5.Parameters.Append prm

set Result5 = oCmd5.Execute


set oCmd51.ActiveConnection = db
oCmd51.CommandText = "armyinformix.dbo.StoredProcedures_4"
oCmd51.CommandType = adCmdStoredProc

iAction = "2"

set prm = oCmd51.CreateParameter("@iAction",adChar,adParamInput,1,iAction)
oCmd51.Parameters.Append prm
set prm = oCmd51.CreateParameter("@sDate",adChar,adParamInput,10,FromDate)
oCmd51.Parameters.Append prm
set prm = oCmd51.CreateParameter("@eDate",adChar,adParamInput,10,ToDate)
oCmd51.Parameters.Append prm

set Result51 = oCmd51.Execute

%>
			<table width="940" height="5" border="0" cellspacing="0" cellpadding="0" align="center"><tr><td></td></tr></table>
			<table  <%=Table_width_and_border%> cellpadding="1" cellspacing="1" bgcolor="#EFECE5" align="center">
			    <tr height="30">
					<td align="left" bgcolor="#FFFFFF" class="TDCont" colspan="6">&nbsp;<%if EXCEL_CHK = "Y" then%>▶<%else%><img src="/Images/dot_01.gif" ><%end if%>&nbsp;<b><font color="#ff00ff"></font> 불만족현황(소속)</b></td>
				</tr>
				<tr bgcolor='#EEF6FF'>
					<td align='center' class='TDCont'  width='150'>구분</td>
					<td align='center' class='TDCont'  width='150'>계</td>
					<td align='center' class='TDCont' width='150'>1군</td>
					<td align='center' class='TDCont' width='150'>2군</td>
					<td align='center' class='TDCont' width='150'>3군</td>
					<td align='center' class='TDCont' width='150'>기타</td>
				</tr>
				<tr bgcolor='#ffffff'>
					<td align='center' class='TDCont' width='150' nowrap bgcolor='#EEF6FF'>건수/(%)</td>
<%
if not Result5.EOF then

per_sect1 = Result5("value6")
per_sect2 = Result5("value7")
per_sect3 = Result5("value8")
per_sect4 = Result5("value9")
per_sect5 = Result5("value10")

'if inStr(per_sect1,".") > 0 then
	per_sect1 = FormatNumber(cdbl(per_sect1),2)
'end if
'if inStr(per_sect2,".") > 0 then
	per_sect2 = FormatNumber(cdbl(per_sect2),2)
'end if
'if inStr(per_sect3,".") > 0 then
	per_sect3 = FormatNumber(cdbl(per_sect3),2)
'end if
'if inStr(per_sect4,".") > 0 then
	per_sect4 = FormatNumber(cdbl(per_sect4),2)
'end if
'if inStr(per_sect5,".") > 0 then
	per_sect5 = FormatNumber(cdbl(per_sect5),2)
'end if

%>
					<td align='center' class='TDCont' width='150'><%=Result5("value1")%>/<%=mark_code1%><%=per_sect1%>%<%=mark_code2%></td>
					<td align='center' class='TDCont' width='150'><%=Result5("value2")%>/<%=mark_code1%><%=per_sect2%>%<%=mark_code2%></td>
					<td align='center' class='TDCont' width='150'><%=Result5("value3")%>/<%=mark_code1%><%=per_sect3%>%<%=mark_code2%></td>
					<td align='center' class='TDCont' width='150'><%=Result5("value4")%>/<%=mark_code1%><%=per_sect4%>%<%=mark_code2%></td>
					<td align='center' class='TDCont' width='150'><%=Result5("value5")%>/<%=mark_code1%><%=per_sect5%>%<%=mark_code2%></td>
<%
end if
%>
				</tr>

<%
set oCmd5.ActiveConnection = nothing
set oCmd5=nothing
set Result5=nothing
%>				

				<tr bgcolor='#ffffff'>
					<td align='center' class='TDCont' width='150' nowrap bgcolor='#EEF6FF'>인원/(%)</td>
<%
if not Result51.EOF then

per_sect1 = Result51("value6")
per_sect2 = Result51("value7")
per_sect3 = Result51("value8")
per_sect4 = Result51("value9")
per_sect5 = Result51("value10")

'if inStr(per_sect1,".") > 0 then
	per_sect1 = FormatNumber(cdbl(per_sect1),2)
'end if
'if inStr(per_sect2,".") > 0 then
	per_sect2 = FormatNumber(cdbl(per_sect2),2)
'end if
'if inStr(per_sect3,".") > 0 then
	per_sect3 = FormatNumber(cdbl(per_sect3),2)
'end if
'if inStr(per_sect4,".") > 0 then
	per_sect4 = FormatNumber(cdbl(per_sect4),2)
'end if
'if inStr(per_sect5,".") > 0 then
	per_sect5 = FormatNumber(cdbl(per_sect5),2)
'end if

%>
					<td align='center' class='TDCont' width='150'><a href="##" onClick="nLink7('0');"><%=Result51("value1")%></a>/<%=mark_code1%><%=per_sect1%>%<%=mark_code2%></td>
					<td align='center' class='TDCont' width='150'><a href="##" onClick="nLink7('1');"><%=Result51("value2")%></a>/<%=mark_code1%><%=per_sect2%>%<%=mark_code2%></td>
					<td align='center' class='TDCont' width='150'><a href="##" onClick="nLink7('2');"><%=Result51("value3")%></a>/<%=mark_code1%><%=per_sect3%>%<%=mark_code2%></td>
					<td align='center' class='TDCont' width='150'><a href="##" onClick="nLink7('3');"><%=Result51("value4")%></a>/<%=mark_code1%><%=per_sect4%>%<%=mark_code2%></td>
					<td align='center' class='TDCont' width='150'><a href="##" onClick="nLink7('9');"><%=Result51("value5")%></a>/<%=mark_code1%><%=per_sect5%>%<%=mark_code2%></td>
<%
end if
%>
				</tr>
		</table>
<%
set oCmd51.ActiveConnection = nothing
set oCmd51=nothing
set Result51=nothing
%>				