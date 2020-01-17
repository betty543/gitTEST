<%
'<!--부대별/응답자별현황 -->

set oCmd1.ActiveConnection = db
oCmd1.CommandText = "armyinformix.dbo.submenu0208_M"
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
oCmd2.CommandText = "armyinformix.dbo.submenu0208_M"
oCmd2.CommandType = adCmdStoredProc

iAction = "2"

set prm = oCmd2.CreateParameter("@iAction",adChar,adParamInput,1,iAction)
oCmd2.Parameters.Append prm
set prm = oCmd2.CreateParameter("@sDate",adChar,adParamInput,10,FromDate)
oCmd2.Parameters.Append prm
set prm = oCmd2.CreateParameter("@eDate",adChar,adParamInput,10,ToDate)
oCmd2.Parameters.Append prm

set Result2 = oCmd2.Execute

TotalCount = Result2("TotalCount")+10
%>

<table width="940" height="5" border="0" cellspacing="0" cellpadding="0" align="center"><tr><td></td></tr></table>
<table <%=Table_width_and_border%> cellspacing="1" cellpadding="0" width="100%" bgcolor="#CCCCCC" align="center">

	<tr height="25" bgcolor="#F3F3F3" align="center">
		<td rowspan=2 colspan=3><b>부대</b></td>
		<td rowspan=2><b>구분</b></td>
		<td rowspan=2 width=80><b>계</b></td>
		<td colspan=3><b>만족도</b></td>
		<td rowspan=2 width=80><b>통화불능</b></td>
		<td rowspan=2 width=80><b>설문거부</b></td>
		<td rowspan=2 width=80><b>미실시</b></td>
	</tr>

	<tr height="25" bgcolor="#F3F3F3" align="center">
		<td width=80><b>만족</b></td>
		<td width=80><b>보통</b></td>
		<td width=80><b>불만족</b></td>
	</tr>

<%
dim pBudae_name_array
redim pBudae_name_array(8)

pBudae_name_array(1) = "1군"
pBudae_name_array(2) = "2군"
pBudae_name_array(3) = "3군"
pBudae_name_array(5) = "육직"
pBudae_name_array(6) = "국방부"
pBudae_name_array(7) = "기타"
pBudae_name_array(0) = "&nbsp;"
pBudae_name_array(4) = "&nbsp;"

TotalCountLine = TotalCount * 8
dim Table_Array_from
Redim Table_Array_from(TotalCountLine,18)

dim Table_Array_from_persect
Redim Table_Array_from_persect(TotalCountLine,7)


for i = 0 to TotalCountLine
	for j = 0 to 17
		Table_Array_from(i,j) = 0
	next
next

for i = 0 to TotalCountLine
	for j = 0 to 6
		Table_Array_from_persect(i,j) = 0
	next
next

dim bdcode1, bdcode2, bdcode3
bdcode1 = ""
bdcode2 = ""
bdcode3 = ""


i = 0
' 0/부대코드1, 1/부대병1, 2/부대코드2, 3/부대명2, 4/부대코드3, 5/부대명3, 
' 6/구분코드, 7/구분명, 8/계, 9/민족, 10/보통, 11/불만족, 12/통화불능, 13/설문거부, 14/미실시
' 15/부대1 rowspan, 16/부대2 rowspan, 17/부대3 rowspan
Do While not Result1.EOF

if i = 0 then
	Table_Array_from(i,15) = 1
	Table_Array_from(i,16) = 1
	Table_Array_from(i,17) = 1
	row1 = 0
	row2 = 0
	row3 = 0
else
	if bdcode1 = Result1("pBudae_code1") then
		Table_Array_from(row1,15) = Table_Array_from(row1,15) + 1
	else
		Table_Array_from(i,15) = 1
		row1 = i
	end if
	if bdcode2 = Result1("pBudae_code2") then
		Table_Array_from(row2,16) = Table_Array_from(row2,16) + 1
	else
		Table_Array_from(i,16) = 1
		row2 = i
	end if
	if bdcode3 = Result1("pBudae_code3") then
		Table_Array_from(row3,17) = Table_Array_from(row3,17) + 1
	else
		Table_Array_from(i,17) = 1
		row3 = i
	end if	
end if

	cur_sect1 = Result1("sect1")
	cur_sect2 = Result1("sect2")
	cur_sect3 = Result1("sect3")
	cur_sect4 = Result1("sect4")
	cur_sect5 = Result1("sect5")
	cur_sect6 = Result1("sect6")
	cur_sect7 = Result1("sect7")
	
	if Result1("code") = "B00" then
		tot_sect1 = cur_sect1
	end if

	if CInt(tot_sect1) = 0 then
		per_sect1 = 0
	else
		per_sect1 = CDBL((cur_sect1/tot_sect1) * 100)
	end if
	if CInt(cur_sect1) = 0 then
		per_sect2 = 0
		per_sect3 = 0
		per_sect4 = 0
		per_sect5 = 0
		per_sect6 = 0
		per_sect7 = 0
	else
		per_sect2 = CDBL((cur_sect2/cur_sect1) * 100)
		per_sect3 = CDBL((cur_sect3/cur_sect1) * 100)
		per_sect4 = CDBL((cur_sect4/cur_sect1) * 100)
		per_sect5 = CDBL((cur_sect5/cur_sect1) * 100)
		per_sect6 = CDBL((cur_sect6/cur_sect1) * 100)
		per_sect7 = CDBL((cur_sect7/cur_sect1) * 100)
	end if
	

	Table_Array_from_persect(i,0) = FormatNumber(cdbl(per_sect1),2)
	Table_Array_from_persect(i,1) = FormatNumber(cdbl(per_sect2),2)
	Table_Array_from_persect(i,2) = FormatNumber(cdbl(per_sect3),2)
	Table_Array_from_persect(i,3) = FormatNumber(cdbl(per_sect4),2)
	Table_Array_from_persect(i,4) = FormatNumber(cdbl(per_sect5),2)
	Table_Array_from_persect(i,5) = FormatNumber(cdbl(per_sect6),2)
	Table_Array_from_persect(i,6) = FormatNumber(cdbl(per_sect7),2)	

	Table_Array_from(i,0) = Result1("pBudae_code1")
	Table_Array_from(i,1) = pBudae_name_array(cint(Result1("pBudae_code1")))
	Table_Array_from(i,2) = Result1("pBudae_code2")
	Table_Array_from(i,3) = Result1("pBudae_name2")
	Table_Array_from(i,4) = Result1("pBudae_code3")
	Table_Array_from(i,5) = Result1("pBudae_name3")
	Table_Array_from(i,6) = Result1("code")
	Table_Array_from(i,7) = Result1("codename")
	Table_Array_from(i,8) = Result1("sect1")
	Table_Array_from(i,9) = Result1("sect2")
	Table_Array_from(i,10) = Result1("sect3")
	Table_Array_from(i,11) = Result1("sect4")
	Table_Array_from(i,12) = Result1("sect5")
	Table_Array_from(i,13) = Result1("sect6")
	Table_Array_from(i,14) = Result1("sect7")

bdcode1 = Result1("pBudae_code1")
bdcode2 = Result1("pBudae_code2")
bdcode3 = Result1("pBudae_code3")

i = i + 1
Result1.MoveNext
Loop
%>	
	
<%
for i = 0 to i-1
%>	

<tr bgcolor="#FFFFFF" align="center">
<%if Table_Array_from(i,15) > 1 then%>
<td rowspan="<%=Table_Array_from(i,15)%>"><strong><%=Table_Array_from(i,1)%></strong></td>
<%end if%>
<%if Table_Array_from(i,16) > 1 then%>
<td rowspan="<%=Table_Array_from(i,16)%>"><strong><%=Table_Array_from(i,3)%></strong></td>
<%end if%>
<%if Table_Array_from(i,17) > 1 then%>
<td rowspan="<%=Table_Array_from(i,17)%>"><strong><%=Table_Array_from(i,5)%></strong></td>
<%end if%>
<td><strong><%=Table_Array_from(i,7)%></strong></td>
<%
if Table_Array_from(i,6) = "B00" then
%>
<td><strong><%=Table_Array_from(i,8)%><br><%=mark_code1%><%=Table_Array_from_persect(i,0)%>%<%=mark_code2%></strong></td>
<td><strong><%=Table_Array_from(i,9)%><br><%=mark_code1%><%=Table_Array_from_persect(i,1)%>%<%=mark_code2%></strong></td>
<td><strong><%=Table_Array_from(i,10)%><br><%=mark_code1%><%=Table_Array_from_persect(i,2)%>%<%=mark_code2%></strong></td>
<td><strong><%=Table_Array_from(i,11)%><br><%=mark_code1%><%=Table_Array_from_persect(i,3)%>%<%=mark_code2%></strong></td>
<td><strong><%=Table_Array_from(i,12)%><br><%=mark_code1%><%=Table_Array_from_persect(i,4)%>%<%=mark_code2%></strong></td>
<td><strong><%=Table_Array_from(i,13)%><br><%=mark_code1%><%=Table_Array_from_persect(i,5)%>%<%=mark_code2%></strong></td>
<td><strong><%=Table_Array_from(i,14)%><br><%=mark_code1%><%=Table_Array_from_persect(i,6)%>%<%=mark_code2%></strong></td>
<%else%>
<td><%=Table_Array_from(i,8)%><br><%=mark_code1%><%=Table_Array_from_persect(i,0)%>%<%=mark_code2%></td>
<td><%=Table_Array_from(i,9)%><br><%=mark_code1%><%=Table_Array_from_persect(i,1)%>%<%=mark_code2%></td>
<td><%=Table_Array_from(i,10)%><br><%=mark_code1%><%=Table_Array_from_persect(i,2)%>%<%=mark_code2%></td>
<td><%=Table_Array_from(i,11)%><br><%=mark_code1%><%=Table_Array_from_persect(i,3)%>%<%=mark_code2%></td>
<td><%=Table_Array_from(i,12)%><br><%=mark_code1%><%=Table_Array_from_persect(i,4)%>%<%=mark_code2%></td>
<td><%=Table_Array_from(i,13)%><br><%=mark_code1%><%=Table_Array_from_persect(i,5)%>%<%=mark_code2%></td>
<td><%=Table_Array_from(i,14)%><br><%=mark_code1%><%=Table_Array_from_persect(i,6)%>%<%=mark_code2%></td>
<%end if%>
</tr>	

<%
Next
%>	
	
</table>

<%
set oCmd1.ActiveConnection = nothing
set oCmd1 = nothing
set Result1 = nothing

set oCmd2.ActiveConnection = nothing
set oCmd2 = nothing
set Result2 = nothing
%>


