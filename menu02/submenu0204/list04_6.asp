<%
'<!--불만족현황 관계자별-->
set oCmd7.ActiveConnection = db
oCmd7.CommandText = "armyinformix.dbo.StoredProcedures_6"
oCmd7.CommandType = adCmdStoredProc

iAction = "1"

set prm = oCmd7.CreateParameter("@iAction",adChar,adParamInput,1,iAction)
oCmd7.Parameters.Append prm
set prm = oCmd7.CreateParameter("@sDate",adChar,adParamInput,10,FromDate)
oCmd7.Parameters.Append prm
set prm = oCmd7.CreateParameter("@eDate",adChar,adParamInput,10,ToDate)
oCmd7.Parameters.Append prm

set Result7 = oCmd7.Execute

i = 0
count_sum1 = 0
Do While not Result7.EOF

	ArrayValue1(i) = Result7("codename")
	ArrayValue2(i) = Result7("count1")
	
	count_sum1 = count_sum1 + Result7("count1")

i = i + 1
Result7.MoveNext
Loop
%>
			<table width="940" height="5" border="0" cellspacing="0" cellpadding="0" align="center"><tr><td></td></tr></table>
			<table  <%=Table_width_and_border%> cellpadding="1" cellspacing="1" bgcolor="#EFECE5" align="center">
			    <tr height="30">
					<td align="left" bgcolor="#FFFFFF" class="TDCont" colspan="<%=i+2%>">&nbsp;<%if EXCEL_CHK = "Y" then%>▶<%else%><img src="/Images/dot_01.gif" ><%end if%>&nbsp;<b><font color="#ff00ff"></font> 불만족현황(관계자별)</b></td>
				</tr>
				<tr bgcolor='#EEF6FF'>
					<td align='center' class='TDCont'  width='150'>구분</td>
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
				<td align='center' class='TDCont' width='150' nowrap bgcolor='#EEF6FF'>인원(%)</td>	

<%if clng(count_sum1) = 0 then%>

				<td align='center' class='TDCont'  width='150'><a href="##" onClick="nLink('30');"><%=count_sum1%></a>/<%=mark_code1%><%if count_sum1 = 0 then%>0.00%<%else%>100.00%<%end if%><%=mark_code2%></td>
<%
for j = 0 to i-1
%>
					<td align='center' class='TDCont' width='150'>0/<%=mark_code1%>0.00%<%=mark_code2%></td>
<%
next 
%>


<%else%>

				<td align='center' class='TDCont'  width='150'><a href="##" onClick="nLink('30');"><%=count_sum1%></a>/<%=mark_code1%><%=FormatNumber(cdbl(count_sum1*100/count_sum1),2)%>%<%=mark_code2%></td>
<%
for j = 0 to i-1

	per_temp = cdbl(ArrayValue2(j))*100 / count_sum1
'	if inStr(per_temp,".") > 0 then
		per_temp_disp = FormatNumber(cdbl(per_temp),2)
'	else
'		per_temp_disp = per_temp
'	end if	
%>
					<td align='center' class='TDCont' width='150'><a href="##" onClick="nLink('3<%=j+1%>');"><%=ArrayValue2(j)%></a>/<%=mark_code1%><%=per_temp_disp%>%<%=mark_code2%></td>
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
	count_sum1 = 0
next
%>	
				</tr>							
			</table>
			
	
<%
set oCmd7.ActiveConnection = nothing
set oCmd7=nothing
set Result7=nothing
%>					