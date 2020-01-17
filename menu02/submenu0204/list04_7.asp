<%
'<!--불만족현황 총괄 -->

set oCmd8.ActiveConnection = db
oCmd8.CommandText = "armyinformix.dbo.submenu0205"
oCmd8.CommandType = adCmdStoredProc

iAction = "1"

set prm = oCmd8.CreateParameter("@iAction",adChar,adParamInput,1,iAction)
oCmd8.Parameters.Append prm
set prm = oCmd8.CreateParameter("@sDate",adChar,adParamInput,10,FromDate)
oCmd8.Parameters.Append prm
set prm = oCmd8.CreateParameter("@eDate",adChar,adParamInput,10,ToDate)
oCmd8.Parameters.Append prm

set Result8 = oCmd8.Execute
%>
<table width="940" height="5" border="0" cellspacing="0" cellpadding="0" align="center"><tr><td></td></tr></table>
<table  <%=Table_width_and_border%> cellpadding="1" cellspacing="1" bgcolor="#EFECE5" align="center">
    <tr height="30">
		<td align="left" bgcolor="#FFFFFF" class="TDCont" colspan="8">&nbsp;<%if EXCEL_CHK = "Y" then%>▶<%else%><img src="/Images/dot_01.gif" ><%end if%>&nbsp;<b><font color="#ff00ff"></font> 응답유형별현황</b></td>
	</tr>

	<tr height="25" bgcolor="#EEF6FF" align="center">
		<td rowspan=2 width='150'><b>구분</b></td>
		<td rowspan=2 width=100 width='150'><b>계</b></td>
		<td colspan=3><b>만족도</b></td>
		<td rowspan=2 width=100><b>통화불능</b></td>
		<td rowspan=2 width=100><b>설문거부</b></td>
		<td rowspan=2 width=100 bgcolor="#ffffff"><b>미실시</b></td>
	</tr>

	<tr height="25" bgcolor="#EEF6FF" align="center">
		<td width=100><b>만족</b></td>
		<td width=100><b>보통</b></td>
		<td width=100><b>불만족</b></td>
	</tr>

<%
dim tot_sect1, tot_sect2, tot_sect3, tot_sect4, tot_sect5, tot_sect6, tot_sect7
dim cur_sect1, cur_sect2, cur_sect3, cur_sect4, cur_sect5, cur_sect6, cur_sect7
dim per_sect1, per_sect2, per_sect3, per_sect4, per_sect5, per_sect6, per_sect7

Do while not Result8.EOF and Result8("code") = "B00"

	cur_sect1 = Result8("sect1")
	cur_sect2 = Result8("sect2")
	cur_sect3 = Result8("sect3")
	cur_sect4 = Result8("sect4")
	cur_sect5 = Result8("sect5")
	cur_sect6 = Result8("sect6")
	cur_sect7 = Result8("sect7")

	tot_sect1 = cur_sect1
	tot_sect2 = cur_sect1-cur_sect7
	tot_sect3 = cur_sect1-cur_sect7
	tot_sect4 = cur_sect1-cur_sect7
	tot_sect5 = cur_sect1-cur_sect7
	tot_sect6 = cur_sect1-cur_sect7
	tot_sect7 = cur_sect1

	
	if CInt(tot_sect1) = 0 then
		per_sect1 = 0
	else
		per_sect1 = CDBL((cur_sect1/tot_sect1) * 100)
		if inStr(per_sect1,".") > 0 then
			per_sect1 = FormatNumber(cdbl(per_sect1),2)
		end if
	end if
	if CInt(tot_sect2) = 0 then
		per_sect2 = 0
	else
		per_sect2 = CDBL((cur_sect2/tot_sect2) * 100)
		if inStr(per_sect2,".") > 0 then
			per_sect2 = FormatNumber(cdbl(per_sect2),2)
		end if
	end if
	if CInt(tot_sect3) = 0 then
		per_sect3 = 0
	else
		per_sect3 = CDBL((cur_sect3/tot_sect3) * 100)
		if inStr(per_sect3,".") > 0 then
			per_sect3 = FormatNumber(cdbl(per_sect3),2)
		end if
	end if
	if CInt(tot_sect4) = 0 then
		per_sect4 = 0
	else
		per_sect4 = CDBL((cur_sect4/tot_sect4) * 100)
		if inStr(per_sect4,".") > 0 then
			per_sect4 = FormatNumber(cdbl(per_sect4),2)
		end if
	end if
	if CInt(tot_sect5) = 0 then
		per_sect5 = 0
	else
		per_sect5 = CDBL((cur_sect5/tot_sect5) * 100)
		if inStr(per_sect5,".") > 0 then
			per_sect5 = FormatNumber(cdbl(per_sect5),2)
		end if
	end if
	if CInt(tot_sect6) = 0 then
		per_sect6 = 0
	else
		per_sect6 = CDBL((cur_sect6/tot_sect6) * 100)
		if inStr(per_sect6,".") > 0 then
			per_sect6 = FormatNumber(cdbl(per_sect6),2)
		end if
	end if
	if CInt(tot_sect7) = 0 then
		per_sect7 = 0
	else
		per_sect7 = CDBL((cur_sect7/tot_sect7) * 100)
		if inStr(per_sect7,".") > 0 then
			per_sect7 = FormatNumber(cdbl(per_sect7),2)
		end if
	end if


per_sect1 = FormatNumber(cdbl(per_sect1),2)
per_sect2 = FormatNumber(cdbl(per_sect2),2)
per_sect3 = FormatNumber(cdbl(per_sect3),2)
per_sect4 = FormatNumber(cdbl(per_sect4),2)
per_sect5 = FormatNumber(cdbl(per_sect5),2)
per_sect6 = FormatNumber(cdbl(per_sect6),2)
per_sect7 = FormatNumber(cdbl(per_sect7),2)	
%>	
	
		<tr bgcolor="#FFFFFF">
			<td align="center" class='TDCont' bgcolor='#EEF6FF'>인원/(%)</td>
			<td align="center" class='TDCont'><a href="##" onClick="nLink('1');"><%=cur_sect1-cur_sect7%></a>/<%=mark_code1%><%=per_sect1%>%<%=mark_code2%></td>
			<td align="center" class='TDCont'><a href="##" onClick="nLink('2');"><%=cur_sect2%></a>/<%=mark_code1%><%=per_sect2%>%<%=mark_code2%></td>
			<td align="center" class='TDCont'><a href="##" onClick="nLink('3');"><%=cur_sect3%></a>/<%=mark_code1%><%=per_sect3%>%<%=mark_code2%></td>
			<td align="center" class='TDCont'><a href="##" onClick="nLink('4');"><%=cur_sect4%></a>/<%=mark_code1%><%=per_sect4%>%<%=mark_code2%></td>
			<td align="center" class='TDCont'><a href="##" onClick="nLink('5');"><%=cur_sect5%></a>/<%=mark_code1%><%=per_sect5%>%<%=mark_code2%></td>
			<td align="center" class='TDCont'><a href="##" onClick="nLink('6');"><%=cur_sect6%></a>/<%=mark_code1%><%=per_sect6%>%<%=mark_code2%></td>
			<td align="center"><a href="##" onClick="nLink('7');"><%=cur_sect7%></a>/<%=mark_code1%><%=per_sect7%>%<%=mark_code2%></td>
		</tr>

<%
Result8.MoveNext
Loop
%>		
		
</table>

<%
set oCmd8.ActiveConnection = nothing
set oCmd8 = nothing
set Result8 = nothing
%>

