<%
'<!--부대별 -->
set oCmd1.ActiveConnection = db
oCmd1.CommandText = "armyinformix.dbo.StoredProcedures_1"
oCmd1.CommandType = adCmdStoredProc

iAction = "1"

set prm = oCmd1.CreateParameter("@iAction",adChar,adParamInput,1,iAction)
oCmd1.Parameters.Append prm
set prm = oCmd1.CreateParameter("@sDate",adChar,adParamInput,10,FromDate)
oCmd1.Parameters.Append prm
set prm = oCmd1.CreateParameter("@eDate",adChar,adParamInput,10,ToDate)
oCmd1.Parameters.Append prm

set Result1 = oCmd1.Execute

set oCmd2.ActiveConnection = db
oCmd2.CommandText = "armyinformix.dbo.StoredProcedures_1"
oCmd2.CommandType = adCmdStoredProc

iAction = "2"

set prm = oCmd2.CreateParameter("@iAction",adChar,adParamInput,1,iAction)
oCmd2.Parameters.Append prm
set prm = oCmd2.CreateParameter("@sDate",adChar,adParamInput,10,FromDate)
oCmd2.Parameters.Append prm
set prm = oCmd2.CreateParameter("@eDate",adChar,adParamInput,10,ToDate)
oCmd2.Parameters.Append prm

set Result2 = oCmd2.Execute


set oCmd21.ActiveConnection = db
oCmd21.CommandText = "armyinformix.dbo.StoredProcedures_1"
oCmd21.CommandType = adCmdStoredProc

iAction = "3"

set prm = oCmd21.CreateParameter("@iAction",adChar,adParamInput,1,iAction)
oCmd21.Parameters.Append prm
set prm = oCmd21.CreateParameter("@sDate",adChar,adParamInput,10,FromDate)
oCmd21.Parameters.Append prm
set prm = oCmd21.CreateParameter("@eDate",adChar,adParamInput,10,ToDate)
oCmd21.Parameters.Append prm

set Result3 = oCmd21.Execute

set oCmd22.ActiveConnection = db
oCmd22.CommandText = "armyinformix.dbo.StoredProcedures_1"
oCmd22.CommandType = adCmdStoredProc

iAction = "4"

set prm = oCmd22.CreateParameter("@iAction",adChar,adParamInput,1,iAction)
oCmd22.Parameters.Append prm
set prm = oCmd22.CreateParameter("@sDate",adChar,adParamInput,10,FromDate)
oCmd22.Parameters.Append prm
set prm = oCmd22.CreateParameter("@eDate",adChar,adParamInput,10,ToDate)
oCmd22.Parameters.Append prm

set Result4 = oCmd22.Execute

%>

			<table width="940" height="5" border="0" cellspacing="0" cellpadding="0" align="center"><tr><td></td></tr></table>
			<table <%=Table_width_and_border%> cellpadding="1" cellspacing="1" bgcolor="#EFECE5" align="center">
			    <tr height="30">
					<td align="left" bgcolor="#FFFFFF" class="TDCont" colspan="6">&nbsp;<%if EXCEL_CHK = "Y" then%>▶<%else%><img src="/Images/dot_01.gif" ><%end if%>&nbsp;<b><font color="#ff00ff"></font> 부대별</b></td>
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
					<td align='center' class='TDCont' width='150' nowrap bgcolor='#EEF6FF'>건수(사건)</td>
<%
if not Result1.EOF then
%>
					<td align='center' class='TDCont' width='150'><%=Result1("value1")%></td>
					<td align='center' class='TDCont' width='150'><%=Result1("value2")%></td>
					<td align='center' class='TDCont' width='150'><%=Result1("value3")%></td>
					<td align='center' class='TDCont' width='150'><%=Result1("value4")%></td>
					<td align='center' class='TDCont' width='150'><%=Result1("value5")%></td>
<%
end if
%>
				</tr>
				<tr bgcolor='#ffffff'>
					<td align='center' class='TDCont' width='150' nowrap bgcolor='#EEF6FF'>관련인원</td>
<%
if not Result2.EOF then
%>					
					<td align='center' class='TDCont' width='150'><%=Result2("value1")%></td>
					<td align='center' class='TDCont' width='150'><%=Result2("value2")%></td>
					<td align='center' class='TDCont' width='150'><%=Result2("value3")%></td>
					<td align='center' class='TDCont' width='150'><%=Result2("value4")%></td>
					<td align='center' class='TDCont' width='150'><%=Result2("value5")%></td>
<%
end if
%>					
				</tr>
				<tr bgcolor='#EEF6FF'>
					<td align='center' class='TDCont' width='150' nowrap bgcolor='#EEF6FF'>실시인원</td>
<%
if not Result3.EOF then
%>					
					<td align='center' class='TDCont' width='150'><%=Result3("value1")%></td>
					<td align='center' class='TDCont' width='150'><%=Result3("value2")%></td>
					<td align='center' class='TDCont' width='150'><%=Result3("value3")%></td>
					<td align='center' class='TDCont' width='150'><%=Result3("value4")%></td>
					<td align='center' class='TDCont' width='150'><%=Result3("value5")%></td>
<%
end if
%>					
				</tr>
				<tr bgcolor='#ffffff'>
					<td align='center' class='TDCont' width='150' nowrap bgcolor='#EEF6FF'>미실시인원</td>
<%
if not Result4.EOF then
%>					
					<td align='center' class='TDCont' width='150'><%=Result4("value1")%></td>
					<td align='center' class='TDCont' width='150'><%=Result4("value2")%></td>
					<td align='center' class='TDCont' width='150'><%=Result4("value3")%></td>
					<td align='center' class='TDCont' width='150'><%=Result4("value4")%></td>
					<td align='center' class='TDCont' width='150'><%=Result4("value5")%></td>
<%
end if
%>					
				</tr>
			</table>
<%
set oCmd1.ActiveConnection = nothing
set oCmd1=nothing
set Result1=nothing

set oCmd2.ActiveConnection = nothing
set oCmd2=nothing
set Result2=nothing
%>		