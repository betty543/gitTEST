<!-- #include virtual="/Include/Adovbs.inc" -->
<!-- #include virtual="/Include/Top.asp" -->
<!-- #include virtual="/Include/PopLayer.asp" -->
<%
dim FromDate, ToDate, QueryYN

QueryYN = request("QueryYN")
FromDate = request("FromDate")
ToDate = request("ToDate")

if FromDate = "" then FromDate = left(Date(),7)&"-01" end If
if ToDate = "" then ToDate=date() end If

dim pageWHERE

pageWHERE = "QueryYN=N&FromDate="&FromDate&"&ToDate="&ToDate

dim oCmd1, oCmd2, iAction, Result1, Result2, prm
Set oCmd1=Server.CreateObject("ADODB.Command")
Set oCmd2=Server.CreateObject("ADODB.Command")

set oCmd1.ActiveConnection = db
oCmd1.CommandText = "armyinformix.dbo.submenu0206"
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
oCmd2.CommandText = "armyinformix.dbo.submenu0206"
oCmd2.CommandType = adCmdStoredProc

iAction = "3"

set prm = oCmd2.CreateParameter("@iAction",adChar,adParamInput,1,iAction)
oCmd2.Parameters.Append prm
set prm = oCmd2.CreateParameter("@sDate",adChar,adParamInput,10,FromDate)
oCmd2.Parameters.Append prm
set prm = oCmd2.CreateParameter("@eDate",adChar,adParamInput,10,ToDate)
oCmd2.Parameters.Append prm

set Result2 = oCmd2.Execute

dim TotalCount

TotalCount = CLng(Result2("TotalCount"))

dim Array1, i
redim Array1(TotalCount,30)

dim pre_codename, cur_codename, pre_code, cur_code
dim pre_budaecode1, cur_budaecode1, pre_budaecode2, cur_budaecode2, pre_budaecode3, cur_budaecode3
dim codename_num, code_num, budaecode1_num, budaecode2_num, budaecode3_num, array_cnt, cur_point9, cur_point8, cur_point7

pre_codename = ""
cur_codename = ""
pre_code = ""
cur_code = ""
pre_budaecode1 = ""
cur_budaecode1 = ""
pre_budaecode2 = ""
cur_budaecode2 = ""
pre_budaecode3 = ""
cur_budaecode3 = ""

codename_num = 0
code_num = 0
budaecode1_num = 0
budaecode2_num = 0
budaecode3_num = 0

array_cnt = 0

Do while not Result1.EOF

	cur_codename = Result1("codename")
	cur_code = Result1("code")
	if isnull(Result1("no")) = false then
		cur_budaecode1 = Result1("budaecode1")
		cur_budaecode2 = Result1("budaecode2")
		cur_budaecode3 = Result1("budaecode3")
		cur_point9 = Result1("point9")
		cur_point8 = Result1("point8")
		cur_point7 = Result1("point7")
	else
		cur_budaecode1 = cstr(array_cnt)
		cur_budaecode2 = cstr(array_cnt)
		cur_budaecode3 = cstr(array_cnt)
		cur_point9 = -1
		cur_point8 = -1
		cur_point7 = -1
	end if

	Array1(array_cnt,0) = cur_codename
	Array1(array_cnt,1) = 0
	Array1(array_cnt,2) = cur_code
	Array1(array_cnt,3) = 0
	Array1(array_cnt,4) = cur_budaecode1
	Array1(array_cnt,5) = 0
	Array1(array_cnt,6) = cur_budaecode2
	Array1(array_cnt,7) = 0
	Array1(array_cnt,8) = cur_budaecode3
	Array1(array_cnt,9) = 0
	Array1(array_cnt,10) = Result1("code_content")
	Array1(array_cnt,11) = Result1("budaename1")
	Array1(array_cnt,12) = Result1("budaename2")
	Array1(array_cnt,13) = Result1("budaename3")
	Array1(array_cnt,14) = 0
	Array1(array_cnt,15) = 0
	Array1(array_cnt,16) = 0
	Array1(array_cnt,17) = 0
	Array1(array_cnt,18) = 0
	Array1(array_cnt,19) = 0
	Array1(array_cnt,20) = 0
	Array1(array_cnt,21) = 0

	Array1(array_cnt,22) = 0

	if cur_codename <> pre_codename then
		codename_num = array_cnt
		Array1(CInt(codename_num),1) = CInt(Array1(CInt(codename_num),1)) + 1

		code_num = array_cnt
		Array1(CInt(code_num),3) = CInt(Array1(CInt(code_num),3)) + 1

		budaecode1_num = array_cnt
		Array1(CInt(budaecode1_num),5) = CInt(Array1(CInt(budaecode1_num),5)) + 1

		budaecode2_num = array_cnt
		Array1(CInt(budaecode2_num),7) = CInt(Array1(CInt(budaecode2_num),7)) + 1

		budaecode3_num = array_cnt
		Array1(CInt(budaecode3_num),9) = CInt(Array1(CInt(budaecode3_num),9)) + 1
		
		if CStr(cur_point9) = "1" then
			Array1(CInt(array_cnt),14) = Array1(CInt(array_cnt),14) + 1
		end if
		if CStr(cur_point8) = "1" then
			Array1(CInt(array_cnt),15) = Array1(CInt(array_cnt),15) + 1
		end if
		if CStr(cur_point7) = "1" then
			Array1(CInt(array_cnt),16) = Array1(CInt(array_cnt),16) + 1
		end if
		if CStr(cur_point9) = "0" and CStr(cur_point8) = "0" and CStr(cur_point7) = "0" then
			Array1(CInt(array_cnt),22) = Array1(CInt(array_cnt),22) + 1
		end if
		Array1(CInt(array_cnt),17) = CInt(Array1(CInt(array_cnt),17)) + 1
		Array1(CInt(codename_num),18) = CInt(Array1(CInt(codename_num),18)) + 1
	else
		Array1(CInt(codename_num),1) = CInt(Array1(CInt(codename_num),1)) + 1

		if cur_code <> pre_code then
			code_num = array_cnt
			Array1(CInt(code_num),3) = CInt(Array1(CInt(code_num),3)) + 1
			
			budaecode1_num = array_cnt
			Array1(CInt(budaecode1_num),5) = CInt(Array1(CInt(budaecode1_num),5)) + 1

			budaecode2_num = array_cnt
			Array1(CInt(budaecode2_num),7) = CInt(Array1(CInt(budaecode2_num),7)) + 1

			budaecode3_num = array_cnt
			Array1(CInt(budaecode3_num),9) = CInt(Array1(CInt(budaecode3_num),9)) + 1

			if CStr(cur_point9) = "1" then
				Array1(CInt(array_cnt),14) = Array1(CInt(array_cnt),14) + 1
			end if
			if CStr(cur_point8) = "1" then
				Array1(CInt(array_cnt),15) = Array1(CInt(array_cnt),15) + 1
			end if
			if CStr(cur_point7) = "1" then
				Array1(CInt(array_cnt),16) = Array1(CInt(array_cnt),16) + 1
			end if
			if CStr(cur_point9) = "0" and CStr(cur_point8) = "0" and CStr(cur_point7) = "0" then
				Array1(CInt(array_cnt),22) = Array1(CInt(array_cnt),22) + 1
			end if
			Array1(CInt(array_cnt),17) = 1
			Array1(CInt(codename_num),18) = CInt(Array1(CInt(codename_num),18)) + 1
		else
			Array1(CInt(code_num),3) = CInt(Array1(CInt(code_num),3)) + 1
			Array1(CInt(array_cnt),17) = CInt(Array1(CInt(array_cnt-1),17)) + 1

			if cur_budaecode1 <> pre_budaecode1 then
				budaecode1_num = array_cnt
				Array1(CInt(budaecode1_num),5) = CInt(Array1(CInt(budaecode1_num),5)) + 1

				budaecode2_num = array_cnt
				Array1(CInt(budaecode2_num),7) = CInt(Array1(CInt(budaecode2_num),7)) + 1

				budaecode3_num = array_cnt
				Array1(CInt(budaecode3_num),9) = CInt(Array1(CInt(budaecode3_num),9)) + 1

				if CStr(cur_point9) = "1" then
					Array1(CInt(array_cnt),14) = Array1(CInt(array_cnt),14) + 1
				end if
				if CStr(cur_point8) = "1" then
					Array1(CInt(array_cnt),15) = Array1(CInt(array_cnt),15) + 1
				end if
				if CStr(cur_point7) = "1" then
					Array1(CInt(array_cnt),16) = Array1(CInt(array_cnt),16) + 1
				end if
				if CStr(cur_point9) = "0" and CStr(cur_point8) = "0" and CStr(cur_point7) = "0" then
					Array1(CInt(array_cnt),22) = Array1(CInt(array_cnt),22) + 1
				end if
			else
				Array1(CInt(budaecode1_num),5) = CInt(Array1(CInt(budaecode1_num),5)) + 1

				if cur_budaecode2 <> pre_budaecode2 then
					budaecode2_num = array_cnt
					Array1(CInt(budaecode2_num),7) = CInt(Array1(CInt(budaecode2_num),7)) + 1

					budaecode3_num = array_cnt
					Array1(CInt(budaecode3_num),9) = CInt(Array1(CInt(budaecode3_num),9)) + 1

					if CStr(cur_point9) = "1" then
						Array1(CInt(array_cnt),14) = Array1(CInt(array_cnt),14) + 1
					end if
					if CStr(cur_point8) = "1" then
						Array1(CInt(array_cnt),15) = Array1(CInt(array_cnt),15) + 1
					end if
					if CStr(cur_point7) = "1" then
						Array1(CInt(array_cnt),16) = Array1(CInt(array_cnt),16) + 1
					end if
					if CStr(cur_point9) = "0" and CStr(cur_point8) = "0" and CStr(cur_point7) = "0" then
						Array1(CInt(array_cnt),22) = Array1(CInt(array_cnt),22) + 1
					end if
				else
					Array1(CInt(budaecode2_num),7) = CInt(Array1(CInt(budaecode2_num),7)) + 1

					if cur_budaecode3 <> pre_budaecode3 then
						budaecode3_num = array_cnt
						Array1(CInt(budaecode3_num),9) = CInt(Array1(CInt(budaecode3_num),9)) + 1

						if CStr(cur_point9) = "1" then
							Array1(CInt(array_cnt),14) = Array1(CInt(array_cnt),14) + 1
						end if
						if CStr(cur_point8) = "1" then
							Array1(CInt(array_cnt),15) = Array1(CInt(array_cnt),15) + 1
						end if
						if CStr(cur_point7) = "1" then
							Array1(CInt(array_cnt),16) = Array1(CInt(array_cnt),16) + 1
						end if
						if CStr(cur_point9) = "0" and CStr(cur_point8) = "0" and CStr(cur_point7) = "0" then
							Array1(CInt(array_cnt),22) = Array1(CInt(array_cnt),22) + 1
						end if
					else
						Array1(CInt(budaecode2_num),7) = CInt(Array1(CInt(budaecode2_num),7)) - 1
						Array1(CInt(budaecode1_num),5) = CInt(Array1(CInt(budaecode1_num),5)) - 1
						Array1(CInt(code_num),3) = CInt(Array1(CInt(code_num),3)) - 1
						Array1(CInt(codename_num),1) = CInt(Array1(CInt(codename_num),1)) - 1
						Array1(CInt(array_cnt),17) = CInt(Array1(CInt(array_cnt),17)) - 1

						if CStr(cur_point9) = "1" then
							Array1(CInt(budaecode3_num),14) = Array1(CInt(budaecode3_num),14) + 1
						end if
						if CStr(cur_point8) = "1" then
							Array1(CInt(budaecode3_num),15) = Array1(CInt(budaecode3_num),15) + 1
						end if
						if CStr(cur_point7) = "1" then
							Array1(CInt(budaecode3_num),16) = Array1(CInt(budaecode3_num),16) + 1
						end if
						if CStr(cur_point9) = "0" and CStr(cur_point8) = "0" and CStr(cur_point7) = "0" then
							Array1(CInt(budaecode3_num),22) = Array1(CInt(budaecode3_num),22) + 1
						end if
					end if
				end if
			end if
		end if
	end if

	pre_codename = cur_codename
	pre_code = cur_code
	pre_budaecode1 = cur_budaecode1
	pre_budaecode2 = cur_budaecode2
	pre_budaecode3 = cur_budaecode3
	array_cnt = array_cnt + 1
Result1.MoveNext
Loop
%>

<script>

	function fn_Search() {

		document.inUpFrm.QueryYN.value = "Y";
		document.inUpFrm.submit();
	}
	
	function fn_Xls() {
		location.href="./list06_Xls.asp?<%=pageWHERE%>"
	}


</script>

<div id='CalendarLayer' style='display:none;'><iframe name='CalendarFrame' src='/include/Calendar.asp' width='172' height='177' border='0' frameborder='0' scrolling='no'></iframe></div>

<table border="0" width="940" cellpadding="0" cellspacing="3" bgcolor="#EFECE5" align="center">
	<tr bgcolor="#FFFFFF">
		<td>
		
			<form name="inUpFrm" method="post" action="./list06.asp" onsubmit="return fn_Search(this);" style="margin:0">
			<input type="hidden" name="QueryYN" value="<%=QueryYN%>">
			<table width="100%" border="0" cellspacing="1" cellpadding="1" align="center" bgcolor="#CCCCCC">

			    <tr>
			        <td width="110" bgcolor="#EFEFEF" class="TDCont">조회기간 :</td>
			        <td  bgcolor="#FFFFFF" colspan=3 width=300>
			        	<input value="<%=FromDate%>" name="FromDate" readonly type="text" size="10" onfocus="setFocusColor(this);" onClick="new CalendarFrame.Calendar(this);" onblur="setOutColor(this);">
				    	~
				    	<input value="<%=ToDate%>" name="ToDate" readonly type="text" size="10" onfocus="setFocusColor(this);" onClick="new CalendarFrame.Calendar(this);" onblur="setOutColor(this);">	
			        </td>
			        <td colspan='2' rowspan="2" bgcolor="#FFFFFF" align="center">
			        	<img src="/Images/Btn/BtnSearch.gif" style="cursor:hand;" onClick="fn_Search();">
			        	<%IF SS_Login_Secgroup="A" Or SS_Login_Secgroup="B" THEN%><br><br><img src="/Images/Btn/BtnExcel.gif" style="cursor:hand;" onClick="fn_Xls();"><%END IF%>
			        </td>
				</tr>

			</table>
			</form>
		</td>
	</tr>
</table>


<table width="940" height="5" border="0" cellspacing="0" cellpadding="0" align="center"><tr><td></td></tr></table>
<table border="0" width="940" cellspacing="1" cellpadding="0" width="100%" bgcolor="#CCCCCC" align="center">
<tr height="25" bgcolor="#F3F3F3" align="center">
<td rowspan=2 width="80"><b>응답자</b></td>
<td rowspan=2><b>질문<br>(설문항목)</b></td>
<td rowspan=2 colspan=3><b>부대</b></td>
<td rowspan=2 width="70"><b>계</b></td>
<td colspan=3><b>만족도</b></td>
<td rowspan=2 width="70"><b>미실시</b></td>
</tr>
<tr height="25" bgcolor="#F3F3F3" align="center">
<td width="60"><b>만족</b></td>
<td width="60"><b>보통</b></td>
<td width="60"><b>불만족</b></td>
</tr>
<%
dim temp_total, temp_tt, temp_t1, temp_t2, temp_t3, temp_t4
dim temp_t1_per, temp_t2_per, temp_t3_per
temp_total = 0
temp_tt = 0
temp_t1 = 0
temp_t2 = 0
temp_t3 = 0
temp_t4 = 0

dim line_total, line_per9, line_per8, line_per7, line_per6

for i = 0 to TotalCount-1
%>
	<%
	if CInt(Array1(i,9)) > 0 then
	%>	
<tr bgcolor="#FFFFFF">
	<%
	if CInt(Array1(i,1)) >= 1 then
	%>
<td align="center" rowspan="<%=CInt(Array1(i,1))+CInt(Array1(i,18))%>"><%=Array1(i,0)%></td>
	<%end if%>
	<%
	if CInt(Array1(i,3)) >= 1 then
		temp_total = Array1(i,3)
		line_total = CInt(Array1(i,14)) + CInt(Array1(i,15)) + CInt(Array1(i,16)) + CInt(Array1(i,22))
	%>	
<td align="left" rowspan="<%=CInt(Array1(i,3))+1%>">&nbsp;<%=Array1(i,10)%></td>
	<%
	end if
	%>
	<%
	if CInt(Array1(i,5)) >= 1 then
	%>	
<td align="center" width="90" rowspan="<%=Array1(i,5)%>"><b><%=Array1(i,11)%></b></td>
	<%
	end if
	%>
	<%
	if CInt(Array1(i,7)) >= 1 then
	%>	
<td align="center" width="90" rowspan="<%=Array1(i,7)%>"><b><%=Array1(i,12)%></b></td>
	<%
	end if
	%>
	<%
	line_total = Array1(i,14)+Array1(i,15)+Array1(i,16)+Array1(i,22)
	if CInt(Array1(i,14)) > 0 then
		line_per9 = cdbl(cdbl(Array1(i,14))/cdbl(line_total)*100)
		if inStr(CStr(line_per9),".") > 0 then
			line_per9 = FormatNumber(cdbl(line_per9),2)
		end if
	else 
		line_per9 = 0
	end if
	if CInt(Array1(i,15)) > 0 then
		line_per8 = cdbl(cdbl(Array1(i,15))/cdbl(line_total)*100)
		if inStr(CStr(line_per8),".") > 0 then
			line_per8 = FormatNumber(cdbl(line_per8),2)
		end if
	else 
		line_per8 = 0
	end if
	if CInt(Array1(i,16)) > 0 then
		line_per7 = cdbl(cdbl(Array1(i,16))/cdbl(line_total)*100)
		if inStr(CStr(line_per7),".") > 0 then
			line_per7 = FormatNumber(cdbl(line_per7),2)
		end if
	else 
		line_per7 = 0
	end if

	if CInt(Array1(i,22)) > 0 then
		line_per6 = cdbl(cdbl(Array1(i,22))/cdbl(line_total)*100)
		if inStr(CStr(line_per6),".") > 0 then
			line_per6 = FormatNumber(cdbl(line_per6),2)
		end if
	else 
		line_per6 = 0
	end if
	temp_t1 = temp_t1 + CInt(Array1(i,14))
	temp_t2 = temp_t2 + CInt(Array1(i,15))
	temp_t3 = temp_t3 + CInt(Array1(i,16))
	temp_t4 = temp_t4 + CInt(Array1(i,22))
	
	line_per6 = FormatNumber(cdbl(line_per6),2)
	line_per7 = FormatNumber(cdbl(line_per7),2)
	line_per8 = FormatNumber(cdbl(line_per8),2)
	line_per9 = FormatNumber(cdbl(line_per9),2)
	
	
	%>	
<td align="center" width="90"><b><%=Array1(i,13)%></b></td>
<td align="center"><%=Array1(i,14)+Array1(i,15)+Array1(i,16)+Array1(i,22)%></td>
<td align="center"><%=Array1(i,14)%><br>[<%=line_per9%>%]</td>
<td align="center"><%=Array1(i,15)%><br>[<%=line_per8%>%]</td>
<td align="center"><%=Array1(i,16)%><br>[<%=line_per7%>%]</td>
<td align="center"><%=Array1(i,22)%><br>[<%=line_per6%>%]</td>
</tr>
	<%
	end if
	%>
	<%
	if CInt(temp_total) = CInt(Array1(i,17)) then
		temp_tt = temp_t1 + temp_t2 + temp_t3 + temp_t4
		
		if CInt(temp_t1) > 0 then
			temp_t1_per = cdbl(cdbl(temp_t1)/cdbl(temp_tt)*100)
			if inStr(CStr(temp_t1_per),".") > 0 then
				temp_t1_per = FormatNumber(cdbl(temp_t1_per),2)
			end if
		else 
			temp_t1_per = 0
		end if
		if CInt(temp_t2) > 0 then
			temp_t2_per = cdbl(cdbl(temp_t2)/cdbl(temp_tt)*100)
			if inStr(CStr(temp_t2_per),".") > 0 then
				temp_t2_per = FormatNumber(cdbl(temp_t2_per),2)
			end if
		else 
			temp_t2_per = 0
		end if
		if CInt(temp_t3) > 0 then
			temp_t3_per = cdbl(cdbl(temp_t3)/cdbl(temp_tt)*100)
			if inStr(CStr(temp_t3_per),".") > 0 then
				temp_t3_per = FormatNumber(cdbl(temp_t3_per),2)
			end if
		else 
			temp_t3_per = 0
		end if
		
		if CInt(temp_t4) > 0 then
			temp_t4_per = cdbl(cdbl(temp_t4)/cdbl(temp_tt)*100)
			if inStr(CStr(temp_t4_per),".") > 0 then
				temp_t4_per = FormatNumber(cdbl(temp_t4_per),2)
			end if
		else 
			temp_t4_per = 0
		end if

		temp_t1_per = FormatNumber(cdbl(temp_t1_per),2)
		temp_t2_per = FormatNumber(cdbl(temp_t2_per),2)
		temp_t3_per = FormatNumber(cdbl(temp_t3_per),2)
		temp_t4_per = FormatNumber(cdbl(temp_t4_per),2)
	%>
<tr bgcolor="#FFFFFF">
<td align="center" colspan="3"><strong>총계</strong></td>
<td align="center"><strong><%=temp_tt%></strong></td>
<td align="center"><strong><%=temp_t1%><br>[<%=temp_t1_per%>%]</strong></td>
<td align="center"><strong><%=temp_t2%><br>[<%=temp_t2_per%>%]</strong></td>
<td align="center"><strong><%=temp_t3%><br>[<%=temp_t3_per%>%]</strong></td>
<td align="center"><strong><%=temp_t4%><br>[<%=temp_t4_per%>%]</strong></td>
</tr>	
	<%
		temp_tt = 0
		temp_t1 = 0
		temp_t2 = 0
		temp_t3 = 0
		temp_t4 = 0
	end if
	%>
<%Next%>
</table>

<%
set oCmd1.ActiveConnection = nothing
set oCmd2.ActiveConnection = nothing
set oCmd1 = nothing
set oCmd2 = nothing
set Result1 = nothing
set Result2 = nothing
set prm = nothing
%>

<!-- #include virtual="/Include/Bottom.asp" -->