<!-- #include virtual="/Include/Top_PopUp.asp" -->


<%

'On Error Resume next

	factnum = Request("factnum")
	if factnum = "" then
		factnum = "0000-00-0000"
	end if
	JOBGB = Request("JOBGB")

		receiptfactnum = Request("receiptfactnum")

		'response.write "사건번호:" & receiptfactnum & "----"

	if JOBGB = "D" then

		'delete
		SQL = "delete from armyinformix.dbo.receiptfact where receiptfactnum = '" & receiptfactnum & "'"
		DB.execute(SQL)

		SQL = "delete from armyinformix.dbo.contactlist where factnum = '" & receiptfactnum & "'"

		DB.execute(SQL)


		SQL = "delete from armyinformix.dbo.factpeople where factnum = '" & receiptfactnum & "'"

		DB.execute(SQL)

		SQL = "delete from armyinformix.dbo.Monitor where factnum = '" & receiptfactnum & "'"

		DB.execute(SQL)

			Response.Write ("<script>alert('정상적으로 삭제되었습니다!');parent.location.reload();</script>")	

	elseif JOBGB = "U" then

		dutyman = Request("dutyman")
		contents = Request("contents")
		nameoffact = Request("nameoffact")
		occurplace = Request("occurplace")
		inputdate = Request("inputdate")
		Date2 = Request("Date2")
		Date3 = Request("Date3")
		receiptkind = Request("receiptkind")
		receiptfactnum1 = Request("receiptfactnum1")		

		if mid(inputdate,3,2) <> mid(receiptfactnum,6,2) then
			receiptfactnum = ""
		end if
		if receiptfactnum = "" then

			receiptfactnum = receiptfactnum1 & "-" & mid(inputdate,3,2) &"-"
			'------------------------------------------------------------------------
				SQL = "select max(right(receiptfactnum,4)) from armyinformix.dbo.receiptfact where left(receiptfactnum,9) = '" & receiptfactnum &"9'"
				SET RsGBN = DB.execute(SQL)
				if isnull(RsGBN(0)) then
					receiptfactnum = receiptfactnum & "9001"
				else
					receiptfactnum = receiptfactnum & RsGBN(0) + 1
				end if

		end if
		'' 수정 또는 등록

		'response.write receiptfactnum

		SQL = "select * from armyinformix.dbo.receiptfact where receiptfactnum = '" & receiptfactnum &"'"
		
		SET RsGBN = DB.execute(SQL)
		if RsGBN.eof = false then
			'수정
			SQL = "	UPDATE	armyinformix.dbo.receiptfact SET	"
			SQL = SQL & "	dutyman = '" & dutyman & "'"
			SQL = SQL & "	,	contents = '" & replace(contents,"'","''") & "'"
			SQL = SQL & "	,	nameoffact = '" & replace(nameoffact,"'","''") & "'"
			SQL = SQL & "	,	occurplace = '" & replace(occurplace,"'","''") & "'"
			SQL = SQL & "	,	inputdate = '" & inputdate & "'"
			SQL = SQL & "	,	Date1 = '" & Date2 & "'"
			SQL = SQL & "	,	Date2 = '" & Date3 & "'"
			SQL = SQL & "	,	receiptkind = '" & receiptkind & "'"

			SQL = SQL & "	where receiptfactnum = '" & receiptfactnum &"'"

			DB.execute(SQL)
			Response.Write ("<script>alert('정상적으로 수정되었습니다!');parent.location.reload();</script>")	
			
		else
			'등록

			SQL = "	insert into armyinformix.dbo.receiptfact ( receiptfactnum, dutyman"
			SQL = SQL & "	,	contents ,	nameoffact"
			SQL = SQL & "	,	occurplace,	inputdate"
			SQL = SQL & "	,	Date1,	Date2"
			SQL = SQL & "	,	receiptkind, filecnt, processgb)"
			SQL = SQL & "	values ( '" & receiptfactnum &"' , '" & dutyman & "'"
			SQL = SQL & "	,	'" & replace(contents,"'","''") & "'"
			SQL = SQL & "	,	'" & replace(nameoffact,"'","''") & "'"
			SQL = SQL & "	,	'" & replace(occurplace,"'","''") & "'"
			SQL = SQL & "	,	'" & inputdate & "'"
			SQL = SQL & "	,	'" & Date2 & "'"
			SQL = SQL & "	,	'" & Date3 & "'"
			SQL = SQL & "	,	'" & receiptkind & "',0,'1')"
			
			
			DB.execute(SQL)
			Response.Write ("<script>alert('정상적으로 등록되었습니다!"&receiptfactnum&"');parent.location.reload();</script>")	

		end if

response.write SQL
		factnum = receiptfactnum
	end if

		'2. 쿼리조건절 셋팅
		pageSize = 10
		pageSector = 10
		if curPage = "" then curPage = 1 end If

		where1 = "FromDate=" & FromDate & "&ToDate=" & ToDate & "&whereCD1=" & whereCD1 & "&whereCD2=" & whereCD2 & "&whereCD3=" & whereCD3 & "&whereCD4=" & whereCD4 &"&QUERYGB="&QUERYGB&"&CLASSNAME="&CLASSNAME&"&whereCD6="&whereCD6&"&whereCD5="&whereCD5&"&whereCD7="&whereCD7
		where2 = "curPage=" & curPage & "&" & where1

		SQL = "select * from armyinformix.dbo.receiptfact where receiptfactnum = '" & factnum &"'"

		SET RsGBN = DB.execute(SQL)
		if RsGBN.eof = false then

			dutyman = RsGBN("dutyman")
			receiptfactnum = RsGBN("receiptfactnum")
			contents = RsGBN("contents")
			nameoffact = RsGBN("nameoffact")
			occurplace = RsGBN("occurplace")

			inputdate = trim(RsGBN("inputdate"))
			Date2 = trim(RsGBN("Date1"))
			Date3 = trim(RsGBN("Date2"))
			receiptkind = RsGBN("receiptkind")
			receiptfactnum1 = left(RsGBN("receiptfactnum"),4)

			SQL = "	select name, class, (select name from armyinformix.dbo.pbudae where auth = unit) as unitname from armyinformix.dbo.user1 where id = '" & RsGBN("dutyman") & "'"
			SET Rs = DB.execute(SQL)

			if Rs.eof = false then
				sName = rs("name")
				sClassName =  rs("class")
				sBudae = rs("unitname")
			else
				sName = ""
				sClassName = ""
				sBudae = ""
			end if

			Rs.close
			set Rs = nothing
		else
			receiptfactnum1 = ""

		end if



%>
<form method="post" action="monitor.asp" name="DetailForm" style="margin:0">
<table border="0" width="940" cellpadding="0" cellspacing="1" bgcolor="#EFECE5" align="center">
	<tr bgcolor="#FFFFFF">
		<td>		
			<input value="<%=JOBGB%>" name="JOBGB" type="hidden" size="30">
			<table width="100%" height="5" border="0" cellspacing="0" cellpadding="0" align="center"><tr><td></td></tr></table>
			<table width="100%" border="0" cellspacing="1" cellpadding="1" align="center" bgcolor="#ffffff">
			    <tr>
					<td align="left" bgcolor="#FFFFFF" class="TDCont"><b><font color="#0000ff">&nbsp;<img src="/Images/dot_01.gif">&nbsp;담당수사관정보</font></b></td>
				</tr>
			</table>
			<table width="100%" border="0" cellspacing="1" cellpadding="1" align="center" bgcolor="#CCCCCC">
			    <tr>
					<td bgcolor="#FFEEF9" class="TDCont" width=100 align='center'>수사관코드</td>
					<td bgcolor="#FFFFFF" width=150>&nbsp;<select name="dutyman" size="1" class="ComboFFFCE7" onchange="pCateSelect('1');">
						<option value="">선택</option>
<%					
							if receiptfactnum1 <> "" then
								SQL = "	select * from armyinformix.dbo.user1 where unit = '" & receiptfactnum1 & "' order by name" '수사관정보
							else
								SQL = "	select * from armyinformix.dbo.user1 order by name" '수사관정보
							end if

							
							SET Rs = DB.execute(SQL)
							do until Rs.eof
									CODE = Rs("id")
									CODENAME = Rs("name") & Rs("class")
								%>

									<%=printSelect("" &CODENAME& "","" &CODE& "","" &dutyman& "")%>
								<%
								Rs.movenext
							loop
%>
						</select></td>
					<td bgcolor="#EEF6FF" class="TDCont" width=100 align='center'>소 속</td>
					<td bgcolor="#FFFFFF" width=150>&nbsp;<%=sBudae%>
					</td>
					<td bgcolor="#EEF6FF" class="TDCont" width=100 align='center'>계 급</td>
					<td bgcolor="#FFFFFF" width=150>&nbsp;<%=sClassName%>
					</td>

					<td bgcolor="#EEF6FF" class="TDCont" width=100 align='center'>성 명</td>
					<td bgcolor="#FFFFFF" width=150>&nbsp;<%=sName%>
					</td>
				</tr>
			</table>

			<table width="100%" height="5" border="0" cellspacing="0" cellpadding="0" align="center"><tr><td></td></tr></table>

			<table width="100%" border="0" cellspacing="1" cellpadding="1" align="center" bgcolor="#ffffff">
			    <tr>
					<td align="left" bgcolor="#FFFFFF" class="TDCont" ><b><font color="#0000ff">&nbsp;<img src="/Images/dot_01.gif" >&nbsp;사건정보</font></b></td>
				</tr>
			</table>
			<table width="100%" border="0" cellspacing="1" cellpadding="1" align="center" bgcolor="#CCCCCC">

			    <tr>
					<td bgcolor="#FFEEF9" class="TDCont" width=100 align='center'>인지일자</td>
					<td bgcolor="#FFFFFF" colspan="3" >&nbsp;<input value="<%=inputdate%>" name="inputdate" type="text" size="10" onfocus="setFocusColor(this);" >
						&nbsp;<img src="/Images/icon_sche.gif" title="달력" style="cursor:hand;" align="absmiddle" name="Date1_CAR" onClick="window.open('/Include/Calendar_view.asp?goMonth='+document.DetailForm.inputdate.value+'&firstcode=&FlightCode=&StartEndGu=S&ControlName=opener.document.DetailForm.inputdate','','width=800,height=300,top=200,left=300,scrollbars=no,status=no');" onblur="setOutColor(this);">
						&nbsp;<img src="/Images/Btn/BtnIconDel.gif" title="삭제" style="cursor:hand;" align="absmiddle"onclick="document.all.inputdate.value=''"></td>
				</tr>
			    <tr>
					<td bgcolor="#FFEEF9" class="TDCont" width=100 align='center'>사건번호</td>
					<td bgcolor="#FFFFFF" width=400>&nbsp;<select name="unit" size="1" class="ComboFFFCE7" disabled> 
						<option value=""></option>
<%					
							if receiptfactnum1 <> "" then
								SQL = "	select * from armyinformix.dbo.user1 where unit = '" & receiptfactnum1 & "' order by name" '수사관정보
							else
								SQL = "	select * from armyinformix.dbo.user1 order by name" '수사관정보
							end if
							SET Rs = DB.execute(SQL)
							do until Rs.eof
									CODE = Rs("unit")
									CODENAME = Rs("unit")
								%>

									<%=printSelect("" &CODENAME& "","" &CODE& "","" &receiptfactnum1& "")%>
								<%
								Rs.movenext
							loop
%>
						</select><input value="<%=receiptfactnum%>" name="receiptfactnum" type="text" size="20" onfocus="setFocusColor(this);" readonly><input value="<%=receiptfactnum1%>" name="receiptfactnum1" type="hidden" size="2" onfocus="setFocusColor(this);" readonly></td>
					<td bgcolor="#FFEEF9" class="TDCont" width=100 align='center'>사건장소</td>
					<td bgcolor="#FFFFFF" width=400>&nbsp;<input value="<%=occurplace%>" name="occurplace" type="text" size="60" onfocus="setFocusColor(this);"></td>

				</tr>
			    <tr>
					<td bgcolor="#FFEEF9" class="TDCont" width=100 align='center'>사 건 명</td>
					<td bgcolor="#FFFFFF" width=300>&nbsp;<input value="<%=nameoffact%>" name="nameoffact" type="text" size="40" onfocus="setFocusColor(this);"></td>

<% if receiptkind = "" then %>					
					<td bgcolor="#FFEEF9" class="TDCont" width=100 align='center'>사건유형</td>
<% else %>
					<td bgcolor="#EEF6FF" class="TDCont" width=100 align='center'>사건유형</td>
<% end if%>
					<td bgcolor="#FFFFFF" width=300>&nbsp;

<%							'==================================================
							SqlCode = "SELECT CODE, CODENAME FROM TB_CODE"
							SqlCode = SqlCode& " WHERE USEYN='Y' AND SYSYN='Y' AND CODEGROUP='B09' AND Code = '8'"
							SqlCode = SqlCode& " ORDER BY CODE"
							set RsCode = db.execute(SqlCode)
						%>
						<select name="receiptkind" size="1" class="ComboFFFCE7">
						<Option value =''>사건유형</option>
							<%
								IF NOT(RsCode.Eof OR RsCode.bof) THEN
									DO until RsCode.EOF
										CODE = RsCode("CODE")
										CODENAME = RsCode("CODENAME")
							%>
							<%=printSelect("" &CODENAME& "","" &CODE& "","" &receiptkind& "")%>
							<%
									RsCode.MoveNext
									LOOP
								END IF
								RsCode.Close
								set RsCode = NOTHING
							%>
						</select>					</td>
				</tr>
			    <tr>
					<td bgcolor="#FFEEF9" class="TDCont" width=100 align='center'>사건개요</td>
					<td bgcolor="#FFFFFF" colspan="3" align='left'><textarea name="contents" style="width:99%; height:60" wrap="soft" class="TextareaInput" ><%=contents%></textarea>
					</td>
				</tr>

			    <tr>
					<td bgcolor="#EEF6FF" class="TDCont" width=100 align='center'>조 사 일</td>
					<td bgcolor="#FFFFFF" width=300>&nbsp;<input value="<%=Date2%>" name="Date2" type="text" size="10" onfocus="setFocusColor(this);" >
						&nbsp;<img src="/Images/icon_sche.gif" title="달력" style="cursor:hand;" align="absmiddle" name="Date2_CAR" onClick="window.open('/Include/Calendar_view.asp?goMonth='+document.DetailForm.Date2.value+'&firstcode=&FlightCode=&StartEndGu=S&ControlName=opener.document.DetailForm.Date2','','width=800,height=300,top=200,left=300,scrollbars=no,status=no');" onblur="setOutColor(this);">
						&nbsp;<img src="/Images/Btn/BtnIconDel.gif" title="삭제" style="cursor:hand;" align="absmiddle"onclick="document.all.Date2.value='';">

					</td>
					<td bgcolor="#EEF6FF" class="TDCont" width=100 align='center'>송 치 일</td>
					<td bgcolor="#FFFFFF" width=300>&nbsp;<input value="<%=Date3%>" name="Date3" type="text" size="10" onfocus="setFocusColor(this);" >
						&nbsp;<img src="/Images/icon_sche.gif" title="달력" style="cursor:hand;" align="absmiddle" name="Date3_CAR" onClick="window.open('/Include/Calendar_view.asp?goMonth='+document.DetailForm.Date3.value+'&firstcode=&FlightCode=&StartEndGu=S&ControlName=opener.document.DetailForm.Date3','','width=800,height=300,top=200,left=300,scrollbars=no,status=no');" onblur="setOutColor(this);">
						&nbsp;<img src="/Images/Btn/BtnIconDel.gif" title="삭제" style="cursor:hand;" align="absmiddle"onclick="document.all.Date3.value='';">

					</td>
				</tr>


			</table>
		</td>
	</tr>
</table>

<!--<table width="940" height="10" border="0" cellspacing="0" cellpadding="0" align="center"><tr><td></td></tr></table>-->
<table width="920" height="5" border="0" cellspacing="0" cellpadding="0" align="center"><tr><td></td></tr></table>
<table width="920" height="30" border="0" cellspacing="0" cellpadding="0" align="center">
	<tr height="5"><td colspan="2"></td></tr>
	<tr height="30">
		<td align="left" height="35">
			<img src="/Images/Btn/BtnDel.gif" style="cursor:hand;" align="absmiddle" onClick="javascript:fn_del(document.DetailForm);">
		</td>
		<td align="right" height="35">
			<img src="/Images/Btn/BtnSubmit.gif" style="cursor:hand;" align="absmiddle" onClick="javascript:fn_inup(document.DetailForm);">
			<img src="/Images/Btn/BtnClose.gif" style="cursor:hand;" align="absmiddle" onClick="javascript:parent.HddnPOPLayer();">
		</td>
	</tr>
</table>

</form>

<script>

	function fn_inup(form)
	{

		if ( DetailForm.dutyman.value == '' )
		{
			alert('수사관을 선택하십시오!');
			return false
		}
		if ( DetailForm.inputdate.value == '' )
		{
			alert('인지일자를 선택하십시오!');
			return false
		}

		if ( DetailForm.receiptfactnum1.value == '' )
		{
			alert('수사관을 선택하십시오!');
			return false
		}
		form.JOBGB.value = "U";
		form.submit();
	}

	function fn_del(form)
	{

		if (confirm('정말로 삭제하시겠습니까?'))
		{
			form.JOBGB.value = "D";
			form.submit();
		}

	}
	function pCateSelect(cn){

	//alert(DetailForm.dutyman.selectedIndex);
		if ( DetailForm.dutyman.value == '' )
			document.DetailForm.unit.value ='';
		else
			document.DetailForm.unit.selectedIndex = DetailForm.dutyman.selectedIndex;


		document.DetailForm.receiptfactnum1.value = document.DetailForm.unit.options[DetailForm.dutyman.selectedIndex].value;

		if (document.DetailForm.unit.value != DetailForm.receiptfactnum1.value )
		{
			DetailForm.receiptfactnum.value = "";
		}

	}

</script>
<!-- #include virtual="/Include/Bottom_PopUp.asp" -->