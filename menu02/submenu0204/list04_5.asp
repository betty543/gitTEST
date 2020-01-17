<%
'<!--불만족현황 유형-->
set oCmd6.ActiveConnection = db
oCmd6.CommandText = "armyinformix.dbo.StoredProcedures_5"
oCmd6.CommandType = adCmdStoredProc

iAction = "1"

set prm = oCmd6.CreateParameter("@iAction",adChar,adParamInput,1,iAction)
oCmd6.Parameters.Append prm
set prm = oCmd6.CreateParameter("@sDate",adChar,adParamInput,10,FromDate)
oCmd6.Parameters.Append prm
set prm = oCmd6.CreateParameter("@eDate",adChar,adParamInput,10,ToDate)
oCmd6.Parameters.Append prm

set Result6 = oCmd6.Execute

i = 0
count_sum1 = 0
count_sum2 = 0
Do While not Result6.EOF
	ArrayValue4(i) = Result6("code")
	ArrayValue1(i) = Result6("codename")
	ArrayValue2(i) = Result6("count1")
	ArrayValue3(i) = Result6("count2")	
	count_sum1 = count_sum1 + Result6("count1")
	count_sum2 = count_sum2 + Result6("count2")

i = i + 1
Result6.MoveNext
Loop
%>
			<table width="940" height="5" border="0" cellspacing="0" cellpadding="0" align="center"><tr><td></td></tr></table>
			<table  <%=Table_width_and_border%> cellpadding="1" cellspacing="1" bgcolor="#EFECE5" align="center">
			    <tr height="30">
					<td align="left" bgcolor="#FFFFFF" class="TDCont" colspan="<%=i+2%>">&nbsp;<%if EXCEL_CHK = "Y" then%>▶<%else%><img src="/Images/dot_01.gif" ><%end if%>&nbsp;<b><font color="#ff00ff"></font> 불만족현황(유형)</b></td>
				</tr>
				<tr bgcolor='#EEF6FF'>
					<td align='center' class='TDCont'  width='120'>구분</td>
					<td align='center' class='TDCont'  width='150'>계</td>
<%
for j = 0 to i-1
%>
					<td align='center' class='TDCont' width='150'><%=ArrayValue1(j)%></td>
<%
next 
%>		
				</tr>

				<tr bgcolor='#ffffff'>
				<td align='center' class='TDCont' width='120' nowrap bgcolor='#EEF6FF'>건수(%)</td>	

<%if clng(count_sum1) = 0 then%>

				<td align='center' class='TDCont'  width='150'><%=count_sum1%>/<%=mark_code1%><%if count_sum1 = 0 then%>0.00%<%else%>100.00%<%end if%><%=mark_code2%></td>
<%
for j = 0 to i-1
%>
					<td align='center' class='TDCont' width='150'>0/<%=mark_code1%>0.00%<%=mark_code2%></td>
<%
next 
%>


<%else%>

				<td align='center' class='TDCont'  width='150'><%=count_sum1%>/<%=mark_code1%><%=FormatNumber(cdbl(count_sum1*100/count_sum1),2)%>%<%=mark_code2%></td>
<%
for j = 0 to i-1

	per_temp = cdbl(ArrayValue2(j))*100 / count_sum1
'	if inStr(per_temp,".") > 0 then
		per_temp_disp = FormatNumber(cdbl(per_temp),2)
'	else
'		per_temp_disp = per_temp
'	end if	
%>
					<td align='center' class='TDCont' width='150'><%=ArrayValue2(j)%>/<%=mark_code1%><%=per_temp_disp%>%<%=mark_code2%></td>
<%
	per_temp = 0
	per_temp_disp = 0
next 
%>

<%end if%>

				</tr>							

				
				<tr bgcolor='#ffffff'>
				<td align='center' class='TDCont' width='120' nowrap bgcolor='#EEF6FF'>인원(%)</td>	

<%if clng(count_sum2) = 0 then%>

				<td align='center' class='TDCont'  width='150'><a href="##" onClick="nLink('21');"><%=count_sum2%></a>/<%=mark_code1%><%if count_sum2 = 0 then%>0.00%<%else%>100.00%<%end if%><%=mark_code2%></td>
<%
for j = 0 to i-1
%>
					<td align='center' class='TDCont' width='150'>0/<%=mark_code1%>0.00%<%=mark_code2%></td>
<%
next 
%>


<%else%>

				<td align='center' class='TDCont'  width='150'><a href="##" onClick="nLink('20');"><%=count_sum2%></a>/<%=mark_code1%><%=FormatNumber(cdbl(count_sum2*100/count_sum2),2)%>%<%=mark_code2%></td>
<%
for j = 0 to i-1

	per_temp = cdbl(ArrayValue3(j))*100 / count_sum2
'	if inStr(per_temp,".") > 0 then
		per_temp_disp = FormatNumber(cdbl(per_temp),2)
'	else
'		per_temp_disp = per_temp
'	end if	
%>
					<td align='center' class='TDCont' width='150'><a href="##" onClick="nLink('2<%=j+1%>');"><%=ArrayValue3(j)%></a>/<%=mark_code1%><%=per_temp_disp%>%<%=mark_code2%></td>
<%
	per_temp = 0
	per_temp_disp = 0
next 
%>

<%end if%>


<%
for j = 0 to i-1
	ArrayValue1(j) = 0
	ArrayValue2(j) = 0
	ArrayValue3(j) = 0
	count_sum1 = 0
	count_sum2 = 0
next
%>	
				</tr>							
			</table>
			
	
<%
set oCmd6.ActiveConnection = nothing
set oCmd6=nothing
set Result6=nothing
%>					