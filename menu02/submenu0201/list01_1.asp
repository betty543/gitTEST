<%
'<!--부대별 종합현황 -->
set oCmd1.ActiveConnection = db
oCmd1.CommandText = "armyinformix.dbo.StoredProcedures_7"
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
oCmd2.CommandText = "armyinformix.dbo.StoredProcedures_7"
oCmd2.CommandType = adCmdStoredProc

iAction = "2"

set prm = oCmd2.CreateParameter("@iAction",adChar,adParamInput,1,iAction)
oCmd2.Parameters.Append prm
set prm = oCmd2.CreateParameter("@sDate",adChar,adParamInput,10,FromDate)
oCmd2.Parameters.Append prm
set prm = oCmd2.CreateParameter("@eDate",adChar,adParamInput,10,ToDate)
oCmd2.Parameters.Append prm

set Result2 = oCmd2.Execute

TotalCount1 = Result2("TotalCount1")
TotalCount2 = Result2("TotalCount2")+ 10

pBudae_name_array(1) = "1군"
pBudae_name_array(2) = "2군"
pBudae_name_array(3) = "3군"
pBudae_name_array(5) = "육직"
pBudae_name_array(6) = "국방부"
pBudae_name_array(7) = "기타"

pBudae_name_array(0) = "&nbsp;"
pBudae_name_array(4) = "&nbsp;"

dim Table_Array_from
Redim Table_Array_from(TotalCount2,12)

for i = 0 to TotalCount2-1
	for j = 0 to 11
		Table_Array_from(i,j) = 0
	next
next

dim bdcode1, bdcode2, bdcode3
bdcode1 = ""
bdcode2 = ""
bdcode3 = ""

dim s_sum, m_sum, b_sum
Redim s_sum(4), m_sum(4), b_sum(4)
for k = 0 to 3
	s_sum(k) = 0
	m_sum(k) = 0
	b_sum(k) = 0
next


i = 0
m = 0
n = 0
z = 0
start_chk = "Y"

' 0:부대명1, 1:부대병2, 2:부대명3, 3:계, 4:만족, 5:보통, 6:불만족
Do While not Result1.EOF

if start_chk = "Y" then

	Table_Array_from(i,0) = pBudae_name_array(cint(Result1("pBudae_code1")))
	Table_Array_from(i,1) = Result1("pBudae_name2")
	Table_Array_from(i,2) = Result1("pBudae_name3")
	Table_Array_from(i,3) = 0
	Table_Array_from(i,9) = Result1("pBudae_code1")
	Table_Array_from(i,10) = Result1("pBudae_code2")
	Table_Array_from(i,11) = Result1("pBudae_code3")

	if cdbl(Result1("monitorpoint")) >= 9 then
		Table_Array_from(i,4) = Table_Array_from(i,4) + 1
		b_sum(1) = b_sum(1) + 1
	end if

	if cdbl(Result1("monitorpoint")) < 9 and cdbl(Result1("monitorpoint")) >= 8 then
		Table_Array_from(i,5) = Table_Array_from(i,5) + 1
		b_sum(2) = b_sum(2) + 1
	end if

	if cdbl(Result1("monitorpoint")) < 8 then
		Table_Array_from(i,6) = Table_Array_from(i,6) + 1
		b_sum(3) = b_sum(3) + 1
	end if

	start_chk = "N"

else

	if bdcode3 <> Result1("pBudae_code3") then
		i = i + 1
	end if

	Table_Array_from(i,0) = pBudae_name_array(cint(Result1("pBudae_code1")))
	Table_Array_from(i,1) = Result1("pBudae_name2")
	Table_Array_from(i,2) = Result1("pBudae_name3")
	Table_Array_from(i,3) = 0
	Table_Array_from(i,9) = Result1("pBudae_code1")
	Table_Array_from(i,10) = Result1("pBudae_code2")
	Table_Array_from(i,11) = Result1("pBudae_code3")

	if cdbl(Result1("monitorpoint")) >= 9 then
		Table_Array_from(i,4) = Table_Array_from(i,4) + 1
		b_sum(1) = b_sum(1) + 1
	end if

	if cdbl(Result1("monitorpoint")) < 9 and cdbl(Result1("monitorpoint")) >= 8 then
		Table_Array_from(i,5) = Table_Array_from(i,5) + 1
		b_sum(2) = b_sum(2) + 1
	end if

	if cdbl(Result1("monitorpoint")) < 8 then
		Table_Array_from(i,6) = Table_Array_from(i,6) + 1
		b_sum(3) = b_sum(3) + 1
	end if


end if



bdcode1 = Result1("pBudae_code1")
bdcode2 = Result1("pBudae_code2")
bdcode3 = Result1("pBudae_code3")

Result1.MoveNext
Loop


for j = 0 to i

	if j = 0 then 
		Table_Array_from(m,7) = 1
		Table_Array_from(n,8) = 1
	else
	
		if Table_Array_from(j-1,9) = Table_Array_from(j,9) then
			Table_Array_from(m,7) = Table_Array_from(m,7) + 1
			Table_Array_from(j,7) = 0
		else
			Table_Array_from(j,7) = 1
			m = j
		end if
	
		if Table_Array_from(j-1,10) = Table_Array_from(j,10) then
			Table_Array_from(n,8) = Table_Array_from(n,8) + 1
			Table_Array_from(j,8) = 0
		else
			Table_Array_from(j,8) = 1
			n = j
		end if
		
		if Table_Array_from(j-1,10) <> Table_Array_from(j,10) then
			if Table_Array_from(j-1,9) = Table_Array_from(j,9) then
				Table_Array_from(z,7) = Table_Array_from(z,7) + 1
			else
				Table_Array_from(z,7) = Table_Array_from(z,7) + 1
				z = j
			end if
		end if
		
	end if
next
Table_Array_from(z,7) = Table_Array_from(z,7) + 1

%>
<table width="940" height="5" border="0" cellspacing="0" cellpadding="0" align="center"><tr><td></td></tr></table>
<table <%=Table_width_and_border%> cellspacing="1" cellpadding="0" width="100%" bgcolor="#CCCCCC" align="center">
	<tr height="30">
		<td align="left" bgcolor="#FFFFFF" class="TDCont" colspan="7">&nbsp;<%if EXCEL_CHK = "Y" then%>▶<%else%><img src="/Images/dot_01.gif" ><%end if%>&nbsp;<b><font color="#ff00ff"></font> 종합현황</b></td>
	</tr>
</table>
<%if EXCEL_CHK = "Y" then%>
	<table width="940" height="5" border="0" cellspacing="0" cellpadding="0" align="center"><tr><td></td></tr></table>
	<table <%=Table_width_and_border%> cellspacing="1" cellpadding="0" width="100%" bgcolor="#CCCCCC" align="center">
		<tr height="30">
			<td align="left" bgcolor="#FFFFFF" class="TDCont" colspan="7">&nbsp;&nbsp;<b><font color="#ff00ff">기간:&nbsp;<%=FromDate%>~<%=ToDate%></font> </b></td>
		</tr>
	</table>
<%end if%>

<table width="940" height="5" border="0" cellspacing="0" cellpadding="0" align="center"><tr><td></td></tr></table>
<table <%=Table_width_and_border%> cellspacing="1" cellpadding="0" width="100%" bgcolor="#CCCCCC" align="center">

	<tr height="25" bgcolor="#F3F3F3" align="center">
		<td rowspan=2 colspan=3><b>부대</b></td>
		<td rowspan=2 width=120><b>계</b></td>
		<td colspan=3><b>만족도</b></td>
	</tr>

	<tr height="25" bgcolor="#F3F3F3" align="center">
		<td width=120><b>만족</b></td>
		<td width=120><b>보통</b></td>
		<td width=120><b>불만족</b></td>
	</tr>
	
		
<%if TotalCount1 > 0 then%>
		
	
		<tr bgcolor="#FFFFFF">
			<td colspan="3" align="center"><b>총계</b></td>
			<td align="center"><strong><%=b_sum(1)+b_sum(2)+b_sum(3)%></strong></td>
			<%
		per_sect = CDBL((b_sum(1)/(b_sum(1)+b_sum(2)+b_sum(3))) * 100)
'		if inStr(per_sect,".") > 0 then
			per_sect = FormatNumber(cdbl(per_sect),2)
'		end if			
			%>
			<td align="center"><strong><%=b_sum(1)%><br><%=mark_code1%><%=per_sect%>%<%=mark_code2%></strong></td>
			<%
		per_sect = CDBL((b_sum(2)/(b_sum(1)+b_sum(2)+b_sum(3))) * 100)
'		if inStr(per_sect,".") > 0 then
			per_sect = FormatNumber(cdbl(per_sect),2)
'		end if			
			%>			
			<td align="center"><strong><%=b_sum(2)%><br><%=mark_code1%><%=per_sect%>%<%=mark_code2%></strong></td>
			<%
		per_sect = CDBL((b_sum(3)/(b_sum(1)+b_sum(2)+b_sum(3))) * 100)
'		if inStr(per_sect,".") > 0 then
			per_sect = FormatNumber(cdbl(per_sect),2)
'		end if			
			%>			
			<td align="center"><strong><%=b_sum(3)%><br><%=mark_code1%><%=per_sect%>%<%=mark_code2%></strong></td>
		</tr>	
<%
per_sect = 0

for j = 0 to i
%>

<%
if j > 0 then
if Table_Array_from(j-1,10) <> Table_Array_from(j,10) then%>		
		<tr bgcolor="#FFFFFF">
			<td align="center"><strong>소계</strong></td>
			<td align="center"><strong><%=s_sum(1)+s_sum(2)+s_sum(3)%></strong></td>
			<%
		per_sect = CDBL((s_sum(1)/(s_sum(1)+s_sum(2)+s_sum(3))) * 100)
'		if inStr(per_sect,".") > 0 then
			per_sect = FormatNumber(cdbl(per_sect),2)
'		end if			
			%>				
			<td align="center"><strong><%=s_sum(1)%><br><%=mark_code1%><%=per_sect%>%<%=mark_code2%></strong></td>
			<%
		per_sect = CDBL((s_sum(2)/(s_sum(1)+s_sum(2)+s_sum(3))) * 100)
'		if inStr(per_sect,".") > 0 then
			per_sect = FormatNumber(cdbl(per_sect),2)
'		end if			
			%>					
			<td align="center"><strong><%=s_sum(2)%><br><%=mark_code1%><%=per_sect%>%<%=mark_code2%></strong></td>
			<%
		per_sect = CDBL((s_sum(3)/(s_sum(1)+s_sum(2)+s_sum(3))) * 100)
'		if inStr(per_sect,".") > 0 then
			per_sect = FormatNumber(cdbl(per_sect),2)
'		end if			
			%>					
			<td align="center"><strong><%=s_sum(3)%><br><%=mark_code1%><%=per_sect%>%<%=mark_code2%></strong></td>
		</tr>
<%
s_sum(1) = 0
s_sum(2) = 0
s_sum(3) = 0
end if%>		
<%if Table_Array_from(j-1,9) <> Table_Array_from(j,9) then%>		
		<tr bgcolor="#FFFFFF">
			<td align="center" colspan="2"><strong>소계</strong></td>
			<td align="center"><strong><%=m_sum(1)+m_sum(2)+m_sum(3)%></strong></td>
			<%
		per_sect = CDBL((m_sum(1)/(m_sum(1)+m_sum(2)+m_sum(3))) * 100)
'		if inStr(per_sect,".") > 0 then
			per_sect = FormatNumber(cdbl(per_sect),2)
'		end if			
			%>					
			<td align="center"><strong><%=m_sum(1)%><br><%=mark_code1%><%=per_sect%>%<%=mark_code2%></strong></td>
			<%
		per_sect = CDBL((m_sum(2)/(m_sum(1)+m_sum(2)+m_sum(3))) * 100)
'		if inStr(per_sect,".") > 0 then
			per_sect = FormatNumber(cdbl(per_sect),2)
'		end if			
			%>					
			<td align="center"><strong><%=m_sum(2)%><br><%=mark_code1%><%=per_sect%>%<%=mark_code2%></strong></td>
			<%
		per_sect = CDBL((m_sum(3)/(m_sum(1)+m_sum(2)+m_sum(3))) * 100)
'		if inStr(per_sect,".") > 0 then
			per_sect = FormatNumber(cdbl(per_sect),2)
'		end if			
			%>					
			<td align="center"><strong><%=m_sum(3)%><br><%=mark_code1%><%=per_sect%>%<%=mark_code2%></strong></td>
		</tr>
<%
m_sum(1) = 0
m_sum(2) = 0
m_sum(3) = 0
end if
end if
%>	
		<tr bgcolor="#FFFFFF">
		<%if Table_Array_from(j,7) > 0 then%>
			<td align="center" rowspan="<%=Table_Array_from(j,7)+1%>"><b><%=Table_Array_from(j,0)%></b></td>
		<%end if%>
		<%if Table_Array_from(j,8) > 0 then%>
			<td class="TDCont" align="center" rowspan="<%=Table_Array_from(j,8)+1%>"><%=Table_Array_from(j,1)%></td>
		<%end if%>
			<td class="TDCont" align="center"><%=Table_Array_from(j,2)%></td>
			<td align="center"><%=Table_Array_from(j,4)+Table_Array_from(j,5)+Table_Array_from(j,6)%></td>
			<%
		per_sect = CDBL((Table_Array_from(j,4)/(Table_Array_from(j,4)+Table_Array_from(j,5)+Table_Array_from(j,6))) * 100)
'		if inStr(per_sect,".") > 0 then
			per_sect = FormatNumber(cdbl(per_sect),2)
'		end if			
			%>				
			<td align="center"><%=Table_Array_from(j,4)%><br><%=mark_code1%><%=per_sect%>%<%=mark_code2%></td>
			<%
		per_sect = CDBL((Table_Array_from(j,5)/(Table_Array_from(j,4)+Table_Array_from(j,5)+Table_Array_from(j,6))) * 100)
'		if inStr(per_sect,".") > 0 then
			per_sect = FormatNumber(cdbl(per_sect),2)
'		end if			
			%>				
			<td align="center"><%=Table_Array_from(j,5)%><br><%=mark_code1%><%=per_sect%>%<%=mark_code2%></td>
			<%
		per_sect = CDBL((Table_Array_from(j,6)/(Table_Array_from(j,4)+Table_Array_from(j,5)+Table_Array_from(j,6))) * 100)
'		if inStr(per_sect,".") > 0 then
			per_sect = FormatNumber(cdbl(per_sect),2)
'		end if			
			%>				
			<td align="center"><%=Table_Array_from(j,6)%><br><%=mark_code1%><%=per_sect%>%<%=mark_code2%></td>
		</tr>	
		
<%
s_sum(1) = s_sum(1) + Table_Array_from(j,4)
s_sum(2) = s_sum(2) + Table_Array_from(j,5)
s_sum(3) = s_sum(3) + Table_Array_from(j,6)

m_sum(1) = m_sum(1) + Table_Array_from(j,4)
m_sum(2) = m_sum(2) + Table_Array_from(j,5)
m_sum(3) = m_sum(3) + Table_Array_from(j,6)
next
%>

<%
	'response.write "탄다" & Table_Array_from(j-1,10)
if Table_Array_from(j-1,10) <> Table_Array_from(j,10) then%>		
		<tr bgcolor="#FFFFFF">
			<td align="center"><strong>소계</strong></td>
			<td align="center"><strong><%=s_sum(1)+s_sum(2)+s_sum(3)%></strong></td>
			<%
		per_sect = CDBL((s_sum(1)/(s_sum(1)+s_sum(2)+s_sum(3))) * 100)
'		if inStr(per_sect,".") > 0 then
			per_sect = FormatNumber(cdbl(per_sect),2)
'		end if			
			%>				
			<td align="center"><strong><%=s_sum(1)%><br><%=mark_code1%><%=per_sect%>%<%=mark_code2%></strong></td>
			<%
		per_sect = CDBL((s_sum(2)/(s_sum(1)+s_sum(2)+s_sum(3))) * 100)
'		if inStr(per_sect,".") > 0 then
			per_sect = FormatNumber(cdbl(per_sect),2)
'		end if			
			%>					
			<td align="center"><strong><%=s_sum(2)%><br><%=mark_code1%><%=per_sect%>%<%=mark_code2%></strong></td>
			<%
		per_sect = CDBL((s_sum(3)/(s_sum(1)+s_sum(2)+s_sum(3))) * 100)
'		if inStr(per_sect,".") > 0 then
			per_sect = FormatNumber(cdbl(per_sect),2)
'		end if			
			%>					
			<td align="center"><strong><%=s_sum(3)%><br><%=mark_code1%><%=per_sect%>%<%=mark_code2%></strong></td>
		</tr>
<%
s_sum(1) = 0
s_sum(2) = 0
s_sum(3) = 0
end if%>		
<%if Table_Array_from(j-1,9) <> Table_Array_from(j,9) then%>		
		<tr bgcolor="#FFFFFF">
			<td align="center" colspan="2"><strong>소계</strong></td>
			<td align="center"><strong><%=m_sum(1)+m_sum(2)+m_sum(3)%></strong></td>
			<%
		per_sect = CDBL((m_sum(1)/(m_sum(1)+m_sum(2)+m_sum(3))) * 100)
'		if inStr(per_sect,".") > 0 then
			per_sect = FormatNumber(cdbl(per_sect),2)
'		end if			
			%>					
			<td align="center"><strong><%=m_sum(1)%><br><%=mark_code1%><%=per_sect%>%<%=mark_code2%></strong></td>
			<%
		per_sect = CDBL((m_sum(2)/(m_sum(1)+m_sum(2)+m_sum(3))) * 100)
'		if inStr(per_sect,".") > 0 then
			per_sect = FormatNumber(cdbl(per_sect),2)
'		end if			
			%>					
			<td align="center"><strong><%=m_sum(2)%><br><%=mark_code1%><%=per_sect%>%<%=mark_code2%></strong></td>
			<%
		per_sect = CDBL((m_sum(3)/(m_sum(1)+m_sum(2)+m_sum(3))) * 100)
'		if inStr(per_sect,".") > 0 then
			per_sect = FormatNumber(cdbl(per_sect),2)
'		end if			
			%>					
			<td align="center"><strong><%=m_sum(3)%><br><%=mark_code1%><%=per_sect%>%<%=mark_code2%></strong></td>
		</tr>
<%
m_sum(1) = 0
m_sum(2) = 0
m_sum(3) = 0
end if
%>	


<%end if%>
		
</table>

<%
set oCmd1.ActiveConnection = nothing
set oCmd1=nothing
set Result1=nothing

set oCmd2.ActiveConnection = nothing
set oCmd2=nothing
set Result2=nothing
%>		