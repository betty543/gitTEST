<%
'<!--부대별 종합현황 -->
set oCmd3.ActiveConnection = db
oCmd3.CommandText = "armyinformix.dbo.StoredProcedures_8"
oCmd3.CommandType = adCmdStoredProc

iAction = "1"

set prm = oCmd3.CreateParameter("@iAction",adChar,adParamInput,1,iAction)
oCmd3.Parameters.Append prm
set prm = oCmd3.CreateParameter("@sDate",adChar,adParamInput,10,FromDate)
oCmd3.Parameters.Append prm
set prm = oCmd3.CreateParameter("@eDate",adChar,adParamInput,10,ToDate)
oCmd3.Parameters.Append prm

set Result3 = oCmd3.Execute

set oCmd4.ActiveConnection = db
oCmd4.CommandText = "armyinformix.dbo.StoredProcedures_8"
oCmd4.CommandType = adCmdStoredProc

iAction = "2"

set prm = oCmd4.CreateParameter("@iAction",adChar,adParamInput,1,iAction)
oCmd4.Parameters.Append prm
set prm = oCmd4.CreateParameter("@sDate",adChar,adParamInput,10,FromDate)
oCmd4.Parameters.Append prm
set prm = oCmd4.CreateParameter("@eDate",adChar,adParamInput,10,ToDate)
oCmd4.Parameters.Append prm

set Result4 = oCmd4.Execute

TotalCount1 = Result4("TotalCount1")



pBudae_name_array(1) = "1군"
pBudae_name_array(2) = "2군"
pBudae_name_array(3) = "3군"
pBudae_name_array(5) = "육직"
pBudae_name_array(6) = "국방부"
pBudae_name_array(7) = "기타"

pBudae_name_array(0) = "&nbsp;"
pBudae_name_array(4) = "&nbsp;"

dim Table_Array_from2
Redim Table_Array_from2(TotalCount1,13)

for i = 0 to TotalCount1-1
	for j = 0 to 12
		Table_Array_from2(i,j) = 0
	next
next

bdcode1 = ""
bdcode2 = ""
bdcode3 = ""

i = 0
m = 0
n = 0
y = 0
z = 0
x = 0

' 0:부대명1, 1:부대병2, 2:부대명3, 3:계, 4:만족, 5:보통, 6:불만족
Do While not Result3.EOF

Table_Array_from2(i,0) = pBudae_name_array(cint(Result3("pBudae_code1")))
Table_Array_from2(i,1) = Result3("pBudae_code1")
Table_Array_from2(i,2) = Result3("pBudae_name2")
Table_Array_from2(i,3) = Result3("pBudae_code2")
Table_Array_from2(i,4) = Result3("pBudae_name3")
Table_Array_from2(i,5) = Result3("pBudae_code3")
Table_Array_from2(i,6) = Result3("class")
Table_Array_from2(i,7) = Result3("name")
Table_Array_from2(i,8) = "[<a href='##' onClick=""nLink('"&Result3("receiptfactnum")&"');"">"&Result3("receiptfactnum")&"</a>] "&Result3("nameoffact")
Table_Array_from2(i,9) = Result3("monitorpoint")
Table_Array_from2(i,10) = 0
Table_Array_from2(i,11) = 0
Table_Array_from2(i,12) = 0

i = i + 1
Result3.MoveNext
Loop

for j = 0 to TotalCount1-1

	if j = 0 then 
		Table_Array_from2(m,10) = 1
		Table_Array_from2(n,11) = 1
		Table_Array_from2(y,12) = 1
	else
	
		if Table_Array_from2(j-1,1) = Table_Array_from2(j,1) then
			Table_Array_from2(m,10) = Table_Array_from2(m,10) + 1
			Table_Array_from2(j,10) = 0
		else
			Table_Array_from2(j,10) = 1
			m = j
		end if
	
		if Table_Array_from2(j-1,3) = Table_Array_from2(j,3) then
			Table_Array_from2(n,11) = Table_Array_from2(n,11) + 1
			Table_Array_from2(j,11) = 0
		else
			Table_Array_from2(j,11) = 1
			n = j
		end if

		if Table_Array_from2(j-1,5) = Table_Array_from2(j,5) then
			Table_Array_from2(y,12) = Table_Array_from2(y,12) + 1
			Table_Array_from2(j,12) = 0
		else
			Table_Array_from2(j,12) = 1
			y = j
		end if
		
	end if
next


m = 0
n = 0
y = 0
for j = 1 to TotalCount1-1


		if Table_Array_from2(j-1,5) <> Table_Array_from2(j,5) then
			Table_Array_from2(m,10) = Table_Array_from2(m,10) + 1
			Table_Array_from2(n,11) = Table_Array_from2(n,11) + 1
			Table_Array_from2(y,12) = Table_Array_from2(y,12) + 1
			y = j
			if Table_Array_from2(j-1,3) <> Table_Array_from2(j,3) then
				Table_Array_from2(m,10) = Table_Array_from2(m,10) + 1
				Table_Array_from2(n,11) = Table_Array_from2(n,11) + 1
				n = j
				if Table_Array_from2(j-1,1) <> Table_Array_from2(j,1) then
					Table_Array_from2(m,10) = Table_Array_from2(m,10) + 1
					m = j
				end if
			end if
		end if
		
'response.write j&"/"&m&"/"&n&"/"&y&"<br>"

next
Table_Array_from2(m,10) = Table_Array_from2(m,10) + 3
Table_Array_from2(n,11) = Table_Array_from2(n,11) + 2
Table_Array_from2(y,12) = Table_Array_from2(y,12) + 1

'Table_Array_from2(z,10) = Table_Array_from2(z,10) + 1


%>

<table width="940" height="5" border="0" cellspacing="0" cellpadding="0" align="center"><tr><td></td></tr></table>
<table <%=Table_width_and_border%> cellspacing="1" cellpadding="0" width="100%" bgcolor="#CCCCCC" align="center">
	<tr height="30">
		<td align="left" bgcolor="#FFFFFF" class="TDCont" colspan="7">&nbsp;<%if EXCEL_CHK = "Y" then%>▶<%else%><img src="/Images/dot_01.gif" ><%end if%>&nbsp;<b><font color="#ff00ff"></font> 세부현황</b></td>
	</tr>
</table>
<table width="940" height="5" border="0" cellspacing="0" cellpadding="0" align="center"><tr><td></td></tr></table>
<table <%=Table_width_and_border%> cellspacing="1" cellpadding="0" width="100%" bgcolor="#CCCCCC" align="center">
	<tr height="25" bgcolor="#F3F3F3" align="center">
		<td rowspan=2 colspan=3 width="320"><b>부대</b></td>
		<td colspan=2><b>수사관</b></td>
		<td rowspan=2 width=400><b>사건명</b></td>
		<td rowspan=2 width=80><b>만족도</b></td>
	</tr>

	<tr height="25" bgcolor="#F3F3F3" align="center">
		<td width=60><b>계급</b></td>
		<td width=60><b>성명</b></td>
	</tr>

	
<%if TotalCount1 > 0 then%>	

	
<%
for j = 0 to TotalCount1-1
%>
<%
if j > 0 then
if Table_Array_from2(j-1,5) <> Table_Array_from2(j,5) then%>		
		<tr bgcolor="#FFFFFF">
		<%
		avg_sect = CDBL(m_sum(1)/s_sum(1))
		avg_sect = Func_Round(avg_sect, 2)
		%>
			<td align="right" colspan="4"><strong>총 <%=s_sum(1)%> 건 / 평균 <%=avg_sect%> 점(<%=MonitorPointCHK(avg_sect)%>)&nbsp;</strong></td>
		</tr>
<%
s_sum(1) = 0
m_sum(1) = 0
end if%>	
<%if Table_Array_from2(j-1,3) <> Table_Array_from2(j,3) then%>		
		<tr bgcolor="#FFFFFF">
		<%
		avg_sect = CDBL(m_sum(2)/s_sum(2))
		avg_sect = Func_Round(avg_sect, 2)
		%>		
			<td align="right" colspan="5"><strong>총 <%=s_sum(2)%> 건 / 평균 <%=avg_sect%> 점(<%=MonitorPointCHK(avg_sect)%>)&nbsp;</strong></td>
		</tr>
<%
s_sum(1) = 0
s_sum(2) = 0
m_sum(1) = 0
m_sum(2) = 0
end if%>		
<%if Table_Array_from2(j-1,1) <> Table_Array_from2(j,1) then%>		
		<tr bgcolor="#FFFFFF">
		<%
		avg_sect = CDBL(m_sum(3)/s_sum(3))
		avg_sect = Func_Round(avg_sect, 2)
		%>			
			<td align="right" colspan="6"><strong>총 <%=s_sum(3)%> 건 / 평균 <%=avg_sect%> 점(<%=MonitorPointCHK(avg_sect)%>)&nbsp;</strong></td>
		</tr>
<%
s_sum(1) = 0
s_sum(2) = 0
s_sum(3) = 0
m_sum(1) = 0
m_sum(2) = 0
m_sum(3) = 0
end if
end if
%>	
		<tr bgcolor="#FFFFFF">
<%if Table_Array_from2(j,10) > 0 then%>
			<td align="center" width="80" rowspan="<%=Table_Array_from2(j,10)%>"><b><%=Table_Array_from2(j,0)%></b></td></td>
<%end if%>
<%if Table_Array_from2(j,11) > 0 then%>
			<td class="TDCont" width="120" align="center" rowspan="<%=Table_Array_from2(j,11)%>"><%=Table_Array_from2(j,2)%></td></td>
<%end if%>
<%if Table_Array_from2(j,12) > 0 then%>			
			<td class="TDCont" width="120" align="center" rowspan="<%=Table_Array_from2(j,12)%>"><%=Table_Array_from2(j,4)%></td></td>
<%end if%>
			<td align="center"><%=Table_Array_from2(j,6)%></td></td>
			<td align="center"><%=Table_Array_from2(j,7)%></td></td>
			<td align="left">&nbsp;<%=Table_Array_from2(j,8)%></td></td>
<%
Table_Array_from_2_j_9 = Table_Array_from2(j,9)
%>
			<td align="center"><%=FormatNumber(cdbl(Table_Array_from_2_j_9),2)%></td></td>


		</tr>
<%
s_sum(1) = s_sum(1) + 1
s_sum(2) = s_sum(2) + 1
s_sum(3) = s_sum(3) + 1

m_sum(1) = m_sum(1) + Table_Array_from2(j,9)
m_sum(2) = m_sum(2) + Table_Array_from2(j,9)
m_sum(3) = m_sum(3) + Table_Array_from2(j,9)

next%>


<%
if Table_Array_from2(j-1,5) <> Table_Array_from2(j,5) then%>		
		<tr bgcolor="#FFFFFF">
		<%
		avg_sect = CDBL(m_sum(1)/s_sum(1))
		avg_sect = Func_Round(avg_sect, 2)
		%>			
			<td align="right" colspan="4"><strong>총 <%=s_sum(1)%> 건 /  평균 <%=avg_sect%> 점(<%=MonitorPointCHK(avg_sect)%>)&nbsp;</strong></td>
		</tr>
<%
s_sum(1) = 0
m_sum(1) = 0
end if%>	
<%if Table_Array_from2(j-1,3) <> Table_Array_from2(j,3) then%>		
		<tr bgcolor="#FFFFFF">
		<%
		avg_sect = CDBL(m_sum(2)/s_sum(2))
		avg_sect = Func_Round(avg_sect, 2)
		%>			
			<td align="right" colspan="5"><strong>총 <%=s_sum(2)%> 건 /  평균 <%=avg_sect%> 점(<%=MonitorPointCHK(avg_sect)%>)&nbsp;</strong></td>
		</tr>
<%
s_sum(1) = 0
s_sum(2) = 0
m_sum(1) = 0
m_sum(2) = 0
end if%>		
<%if Table_Array_from2(j-1,1) <> Table_Array_from2(j,1) then%>		
		<tr bgcolor="#FFFFFF">
		<%
		avg_sect = CDBL(m_sum(3)/s_sum(3))
		avg_sect = Func_Round(avg_sect, 2)
		%>			
			<td align="right" colspan="6"><strong>총 <%=s_sum(3)%> 건 /  평균 <%=avg_sect%> 점(<%=MonitorPointCHK(avg_sect)%>)&nbsp;</strong></td>
		</tr>
<%
s_sum(1) = 0
s_sum(2) = 0
s_sum(3) = 0
m_sum(1) = 0
m_sum(2) = 0
m_sum(3) = 0
end if
%>	


<%end if%>

</table>




<%
set oCmd3.ActiveConnection = nothing
set oCmd3=nothing
set Result3=nothing

set oCmd4.ActiveConnection = nothing
set oCmd4=nothing
set Result4=nothing
%>		