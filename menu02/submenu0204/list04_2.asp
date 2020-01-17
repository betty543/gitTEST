<%
'<!--유형별-->
set oCmd3.ActiveConnection = db
oCmd3.CommandText = "armyinformix.dbo.StoredProcedures_2"
oCmd3.CommandType = adCmdStoredProc

iAction = "1"

set prm = oCmd3.CreateParameter("@iAction",adChar,adParamInput,1,iAction)
oCmd3.Parameters.Append prm
set prm = oCmd3.CreateParameter("@sDate",adChar,adParamInput,10,FromDate)
oCmd3.Parameters.Append prm
set prm = oCmd3.CreateParameter("@eDate",adChar,adParamInput,10,ToDate)
oCmd3.Parameters.Append prm

set Result3 = oCmd3.Execute


i = 0
count_sum1 = 0
count_sum2 = 0
count_sum3 = 0
count_sum4 = 0
Do While not Result3.EOF

	ArrayValue1(i) = Result3("codename")
	ArrayValue2(i) = Result3("count1")
	ArrayValue3(i) = Result3("count2")
	ArrayValue4(i) = Result3("count3")
	ArrayValue5(i) = Result3("count4")
	
	count_sum1 = count_sum1 + Result3("count1")
	count_sum2 = count_sum2 + Result3("count2")
	count_sum3 = count_sum3 + Result3("count3")
	count_sum4 = count_sum4 + Result3("count4")

i = i + 1
Result3.MoveNext
Loop
%>
			<table width="940" height="5" border="0" cellspacing="0" cellpadding="0" align="center"><tr><td></td></tr></table>
			<table  <%=Table_width_and_border%> cellpadding="1" cellspacing="1" bgcolor="#EFECE5" align="center">
			    <tr height="30">
					<td align="left" bgcolor="#FFFFFF" class="TDCont" colspan="<%=i+2%>">&nbsp;<%if EXCEL_CHK = "Y" then%>▶<%else%><img src="/Images/dot_01.gif" ><%end if%>&nbsp;<b><font color="#ff00ff"></font> 유형별</b></td>
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
				<td align='center' class='TDCont' width='150' nowrap bgcolor='#EEF6FF'>건수(사건)</td>	
				<td align='center' class='TDCont'  width='150'><%=count_sum1%></td>
<%
for j = 0 to i-1
%>
					<td align='center' class='TDCont' width='150'><%=ArrayValue2(j)%></td>
<%
next 
%>	
				</tr>		
				<tr bgcolor='#ffffff'>
					<td align='center' class='TDCont' width='150' nowrap bgcolor='#EEF6FF'>관련인원</td>
					<td align='center' class='TDCont'  width='150'><%=count_sum2%></td>
<%
for j = 0 to i-1
%>
					<td align='center' class='TDCont' width='150'><%=ArrayValue3(j)%></td>
<%
next 
%>
				<tr bgcolor='#EEF6FF'>
					<td align='center' class='TDCont' width='150' nowrap bgcolor='#EEF6FF'>실시인원</td>
					<td align='center' class='TDCont'  width='150'><%=count_sum3%></td>
<%
for j = 0 to i-1
%>
					<td align='center' class='TDCont' width='150'><%=ArrayValue4(j)%></td>
<%
next 
%>
				<tr bgcolor='#ffffff'>
					<td align='center' class='TDCont' width='150' nowrap bgcolor='#EEF6FF'>미실시인원</td>
					<td align='center' class='TDCont'  width='150'><%=count_sum4%></td>
<%
for j = 0 to i-1
%>
					<td align='center' class='TDCont' width='150'><%=ArrayValue5(j)%></td>
<%
next 

for j = 0 to i-1
	ArrayValue1(j) = 0
	ArrayValue2(j) = 0
	ArrayValue3(j) = 0
	ArrayValue4(j) = 0
	ArrayValue5(j) = 0
	count_sum1 = 0
	count_sum2 = 0
	count_sum3 = 0
	count_sum4 = 0
next
%>	
				</tr>							
			</table>
<%
set oCmd3.ActiveConnection = nothing
set oCmd3=nothing
set Result3=nothing
%>					