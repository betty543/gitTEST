<%
'<!--사건관계자-->
set oCmd4.ActiveConnection = db
oCmd4.CommandText = "armyinformix.dbo.StoredProcedures_3"
oCmd4.CommandType = adCmdStoredProc

iAction = "1"

set prm = oCmd4.CreateParameter("@iAction",adChar,adParamInput,1,iAction)
oCmd4.Parameters.Append prm
set prm = oCmd4.CreateParameter("@sDate",adChar,adParamInput,10,FromDate)
oCmd4.Parameters.Append prm
set prm = oCmd4.CreateParameter("@eDate",adChar,adParamInput,10,ToDate)
oCmd4.Parameters.Append prm

set Result4 = oCmd4.Execute

i = 0
count_sum1 = 0
count_sum2 = 0
count_sum3 = 0

Do While not Result4.EOF

	ArrayValue1(i) = Result4("codename")
	ArrayValue2(i) = Result4("count1")
	ArrayValue3(i) = Result4("count2")
	ArrayValue4(i) = Result4("count3")
	
	count_sum1 = count_sum1 + Result4("count1")
	count_sum2 = count_sum2 + Result4("count2")
	count_sum3 = count_sum3 + Result4("count3")

i = i + 1
Result4.MoveNext
Loop
%>
			<table width="940" height="5" border="0" cellspacing="0" cellpadding="0" align="center"><tr><td></td></tr></table>
			<table  <%=Table_width_and_border%> cellpadding="1" cellspacing="1" bgcolor="#EFECE5" align="center">
			    <tr height="30">
					<td align="left" bgcolor="#FFFFFF" class="TDCont" colspan="<%=i+2%>">&nbsp;<%if EXCEL_CHK = "Y" then%>▶<%else%><img src="/Images/dot_01.gif" ><%end if%>&nbsp;<b><font color="#ff00ff"></font> 사건관계자</b></td>
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
				<td align='center' class='TDCont' width='150' nowrap bgcolor='#EEF6FF'>관련인원</td>	
				<td align='center' class='TDCont'  width='150'><%=count_sum1%></td>
<%
for j = 0 to i-1
%>
					<td align='center' class='TDCont' width='150'><%=ArrayValue2(j)%></td>
<%
next 

%>	
				</tr>		

				<tr bgcolor='#EEF6FF'>
				<td align='center' class='TDCont' width='150' nowrap bgcolor='#EEF6FF'>실시인원</td>	
				<td align='center' class='TDCont'  width='150'><%=count_sum2%></td>
<%
for j = 0 to i-1
%>
					<td align='center' class='TDCont' width='150'><%=ArrayValue3(j)%></td>
<%
next 
%>	
				</tr>	

				<tr bgcolor='#ffffff'>
				<td align='center' class='TDCont' width='150' nowrap bgcolor='#EEF6FF'>미실시인원</td>	
				<td align='center' class='TDCont'  width='150'><%=count_sum3%></td>
<%
for j = 0 to i-1
%>
					<td align='center' class='TDCont' width='150'><%=ArrayValue4(j)%></td>
<%
next 
%>	
				</tr>	
<%
for j = 0 to i-1
	ArrayValue1(j) = 0
	ArrayValue2(j) = 0
	ArrayValue3(j) = 0
	ArrayValue4(j) = 0
	count_sum1 = 0
	count_sum2 = 0
	count_sum3 = 0
next
%>	


			</table>
<%
set oCmd4.ActiveConnection = nothing
set oCmd4=nothing
set Result4=nothing
%>		