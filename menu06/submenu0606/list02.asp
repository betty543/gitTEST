<!-- #include virtual="/Include/Top.asp" -->
<%
	'####### 파라미터 ##################################################################################
	QueryYN = request("QueryYN")
	FromDate = request("FromDate")
	ToDate = request("ToDate")
	whereCD1 = Trim(request("whereCD1"))
	whereCD2 = Trim(request("whereCD2"))
	TelNo = Trim(request("TelNo"))

	SS_Login_Secgroup = SESSION("SS_Login_Secgroup")
	SS_Login_Grade = SESSION("SS_Login_Grade")
	SS_Login_CTIID = SESSION("SS_Login_CTIID")
	SS_Login_EXTNO = SESSION("SS_Login_EXTNO")
	SS_LoginID = SESSION("SS_LoginID")


	If QueryYN = "" Then
		whereCD1 = ""
	End if


	if FromDate = "" then FromDate = date() end If
	if ToDate = "" then ToDate=date() end If

	pageWHERE = "QueryYN="&QueryYN&"&FromDate="&FromDate&"&ToDate="&ToDate&"&whereCD1="&whereCD1&"&whereCD2="&whereCD2&"&TelNo="&TelNo

%>

<div id='CalendarLayer' style='display:none;'><iframe name='CalendarFrame' src='/include/Calendar.asp' width='172' height='177' border='0' frameborder='0' scrolling='no'></iframe></div>

<script>

	function fn_Search() {

		document.inUpFrm.QueryYN.value = "Y";
		document.inUpFrm.submit();
	}
	
	function fn_Xls() {
		location.href="list02_Xls.asp?<%=pageWHERE%>"
	}
</script>
<table border="0" width="1200" cellpadding="0" cellspacing="3" bgcolor="#EFECE5" align="center">
	<tr bgcolor="#FFFFFF">
		<td>
			<form name="inUpFrm" method="post" action="<%=Menu_2nd%>" onsubmit="return fn_Search(this);" style="margin:0">
			<input type="hidden" name="QueryYN" value="<%=QueryYN%>">
			<table width="100%" border="0" cellspacing="1" cellpadding="0" style="border:#E1DED6 solid 1px">
			    <tr>
			        <td class="TDCont">조회기간 :
			        	<input value="<%=FromDate%>" name="FromDate" readonly type="text" size="10" onfocus="setFocusColor(this); new CalendarFrame.Calendar(this);" onblur="setOutColor(this);">
				    	~
				    	<input value="<%=ToDate%>" name="ToDate" readonly type="text" size="10" onfocus="setFocusColor(this); new CalendarFrame.Calendar(this);" onblur="setOutColor(this);">
			        </td>

			        <td class="TDCont">회선구분 :


						<input type="radio" name="whereCD2" value="" class="none" <% if whereCD2 ="" then%> checked <%end if%>> 전체
						<input type="radio" name="whereCD2" value="1" class="none" <% if whereCD2 ="1" then%> checked <%end if%> > 군
						<input type="radio" name="whereCD2" value="2" class="none" <% if whereCD2 ="2" then%> checked <%end if%>> 일반

			        </td>

			        <td class="TDCont">통화구분 :

						<input type="radio" name="whereCD1" value="" class="none" <% if whereCD1 ="" then%> checked <%end if%>> 전체
						<input type="radio" name="whereCD1" value="1" class="none" <% if whereCD1 ="1" then%> checked <%end if%> > 인
						<input type="radio" name="whereCD1" value="2" class="none" <% if whereCD1 ="2" then%> checked <%end if%>> 아웃
			        </td>


			        <td class="TDCont">전화번호:

						<input value="<%=TelNo%>" name="TelNo"  type="text" size="15" onfocus="setFocusColor(this);" onblur="setOutColor(this);">
			        </td>

			        <td class="TDR5px">
			        	<img src="/Images/Btn/BtnSearch.gif" align="absmiddle" style="cursor:hand;" onClick="fn_Search();">
			        	<img src="/Images/Btn/BtnExcel.gif" align="absmiddle" style="cursor:hand;" onClick="fn_Xls();">
			        </td>
			    </tr>
			</table>
			</form>
		</td>
	</tr>
</table>
<table border="0" width="100%" cellpadding="0" cellspacing="0" align="center"><tr height="5"><td></td></tr></table>


<table border="0" cellpadding="0" cellspacing="0" align="center">
	<tr>
		<td>
			<!--<DIV style="OVERFLOW-Y:auto; OVERFLOW-X:auto; MARGIN: 0px 0px 0px 0px; WIDTH:940; HEIGHT:500;">-->
			<table width="1200"  border="0" cellpadding="1" cellspacing="1" bgcolor="#EFECE5" align="center">

				<tr bgcolor='#EEF6FF'>
					<td align='center' class='TDCont'  width='30'>No</td>

				<td align='center' class='TDCont'  width='130'>통화일시</td>
				<td align='center' class='TDCont'  width='130'>통화종료일시</td>
				<td align='center' class='TDCont'  width='100'>통화시간</td>
				<td align='center' class='TDCont'  width='130'>내선번호</td>
				<td align='center' class='TDCont'  width='130'>통화구분</td>
				<td align='center' class='TDCont'  width='130'>회선구분</td>
				<td align='center' class='TDCont'  width='130'>회선번호</td>
				<td align='center' class='TDCont'  width='130'>전화번호</td>


				</tr>
<%

	If QueryYN = "Y" Then

		'사용자명, 구분(일반(인,아웃) - 군(인,아웃))
		SQL = "	select count(*) cnt"
		SQL = SQL & "	from i3_ic.dbo.calldetail where convert(char(10),initiateddate,121) >= '"&FromDate&"' and convert(char(10),initiateddate,121) <= '"&ToDate&"'"
		if whereCD2 = "2" then	'일반회선
			SQL = SQL & "	and	( dnis like 'sip:1001%' or dnis like 'sip:1002%')"	
		elseif whereCD2 = "1" then
			SQL = SQL & "	and	( dnis like 'sip:5001%' or dnis like 'sip:5002%')"				
		end if
		if whereCD1 = "2" then '아웃바운드
			SQL = SQL & "	and	calldirection = 'Outbound'"	
		elseif whereCD1 = "1" then
			SQL = SQL & "	and	calldirection = 'Inbound'"				
		end if
		SQL = SQL & "	and	Calltype in ('External','External Party')"
		SQL = SQL & "	and	Stationid <> 'System'"
		if TelNo <> "" then
			SQL = SQL & "	and	remotenumber like '%" & TelNo & "%'"
		end if



'-----------------------------------------------------------------------
'군전화전체
'-----------------------------------------------------------------------
		set rs = db.execute(SQL)
		i = rs("cnt")


		'사용자명, 구분(일반(인,아웃) - 군(인,아웃))
		SQL = "	select '1' gubun, Stationid, LineId, remotenumber, convert(varchar(19),initiateddate,121) sdate, convert(varchar(19),terminateddate,121) edate, calldirection, calldurationseconds, remotenumbercallid, dnis"
		SQL = SQL & "	from i3_ic.dbo.calldetail where convert(char(10),initiateddate,121) >= '"&FromDate&"' and convert(char(10),initiateddate,121) <= '"&ToDate&"'"
		if whereCD2 = "2" then	'일반회선
			SQL = SQL & "	and	( dnis like 'sip:1001%' or dnis like 'sip:1002%')"	
		elseif whereCD2 = "1" then
			SQL = SQL & "	and	( dnis like 'sip:5001%' or dnis like 'sip:5002%')"				
		end if
		if whereCD1 = "2" then '아웃바운드
			SQL = SQL & "	and	calldirection = 'Outbound'"	
		elseif whereCD1 = "1" then
			SQL = SQL & "	and	calldirection = 'Inbound'"				
		end if
		SQL = SQL & "	and	Calltype in ('External','External Party')"
		SQL = SQL & "	and	Stationid <> 'System'"
		if TelNo <> "" then
			SQL = SQL & "	and	remotenumber like '%" & TelNo & "%'"
		end if

		SQL = SQL & "	order by initiateddate desc "

'-----------------------------------------------------------------------
'군전화전체
'-----------------------------------------------------------------------
		set rs = db.execute(SQL)

		do until rs.eof

			if rs("calldirection") = "Inbound" then
				calldirection = "인"
			else
				calldirection = "아웃"
			end if
			if instr(rs("dnis"),"sip:1001")>0 or  instr(rs("dnis"),"sip:1002")>0 then
				LineId = "일반"
				bgcolor = "#EEF6FF"
			elseif instr(rs("dnis"),"sip:5001")>0 or  instr(rs("dnis"),"sip:5002")>0 then
				LineId = "군전화"
				bgcolor = "#ffffff"

			else
				LineId = "일반"
				bgcolor = "#ffffff"
			end if'sip:1832@16.1.153.6:5060

			lv_CallTime = rs("calldurationseconds")
			lv_Hour = Fix(lv_CallTime / 3600)
			lv_Min = Fix((lv_CallTime - lv_Hour * 3600) / 60)
			lv_Sec = lv_CallTime - ((lv_Hour * 3600) + (lv_Min * 60))

			if lv_Hour < 10 then
				lv_Hour = "0" & lv_Hour
			end if
			if lv_Min < 10 then
				lv_Min = "0" & lv_Min
			end if
			if lv_Sec < 10 then
				lv_Sec = "0" & lv_Sec
			end if

			remotenumber = rs("remotenumber")
			cid = replace(replace(replace(replace(remotenumber,"sip:",""),"@1.1.160.85:5060",""),"@1.1.160.89:5060",""),"@audiocodes.com:5060","")
			if instr(cid,"anonymous")>0 then
				cid = "" 
			end if
			if len(cid) = 9 and left(cid,1) <> "0" then
				cid = "0" & cid
			end if


%>
				<tr bgcolor='<%=bgcolor%>'>
				<td align='center'><%=i%></td>
				<td align='center'><%=rs("sdate")%></td>
				<td align='center'><%=rs("edate")%></td>
				<td align='center'><%if lv_Hour <> "00" then%><%=lv_Hour%>:<%end if%><%=lv_Min%>:<%=lv_Sec%></td>

				<td align='center'><%=right(rs("Stationid"),4)%></td>
				<td align='center'><%=calldirection%></td>
				<td align='center'><%=LineId%></td>
				<td align='center'><%=mid(rs("dnis"),5,4)%></td>
				<td align='center'><%=FormatTELNo(cid)%></td></tr>	

<%
			
i = i - 1
			rs.movenext
		loop
%>

	<% End if %>

			</table>
		</td>
	</tr>
</table>



<!-- #include virtual="/Include/Bottom.asp" -->