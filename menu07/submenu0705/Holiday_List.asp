<!-- #include virtual="/Include/Top_Frame.asp" -->
<%
	'1. 파라미터 얻어오기
	sGijunYear = Request("GijunYear")
	FRM = Request("FRM")
	curPage = Request("curPage")

	guboon = Request("guboon")
	sHoliday = Request("Holiday")

	if guboon = "DEL" then
		'내년것도 반영한다.
		sql = "	delete	from	T_Holiday	where [Holiday] = '" & sHoliday & "'"
		db.execute(sql)

		sql = "	delete	from	T_Holiday	where [Holiday] = '" & dateadd("yyyy",1,sHoliday) & "' and Every_Year = '1'"		
		db.execute(sql)

	end if

	'매년적용
	sql = "	select	*	from	T_Holiday where	left([Holiday],4) = '" & sGijunYear & "'"
	sql = sql & "	order by [Holiday]"

	sql_tb = "T_Holiday"
	sql_where = "left([Holiday],4) = '" & sGijunYear & "'"

	'3. 쿼리 실행
	'sql = db_getSqlWithPage(sql_tb, sql_index, sql_field, sql_where, sql_orderby, pageSize, curPage)
	set rs = db.execute(sql)
	'totalCount = rs.recordcount

	'Response.Write sql

	'4. Paging HTML 작성
	totalCount = db_getCount(db, sql_tb, sql_where)
	'startRow = totalCount - pageSize * (curPage - 1)
	'pageHtml = getPageHtml(pageSector, pageSize, totalCount, curPage, currentURL & "?" & where1)

%>

<form method="post" name="inUpFrm" action="<%=currentURL%>" style="margin:0">
<input type="hidden" name="guboon" value="">
<input type="hidden" name="Holiday" value="">
<input type="hidden" name="GijunYear" value="<%=sGijunYear%>">
</form>
<table width="940" height="5" border="0" cellspacing="0" cellpadding="0" align="center"><tr><td></td></tr></table>
<table width="940" border="0" cellspacing="1" cellpadding="1" align="center" bgcolor="#CCCCCC">
	<tr height="20" bgcolor="#EEF6FF" align="center">
		<td width="50">No</td>
		<td width="80">일자</td>
		<td width="150">휴일명</td>
		<td width="80">매년적용</td>
		<td width="120">작업</td>
		<td>비고</td>
	</tr>
	<tr><td colspan="6" height="1" bgcolor="#FFFFFF"></td></tr>
<% IF (RS.EOF OR RS.BOF) THEN %>
	<tr><td height="50" colspan="50" align="center" bgcolor="#FFFFFF" style="color:#0000FF"><%=sGijunYear%>해당년도에 검색된 자료가 없습니다.</td></tr>
<%
	ELSE

		iii = 0

		do until rs.EOF

			iii = iii + 1
			tmpBgColor = "#FFFFFF"

			sHoliday = rs("Holiday")
			sHoliday_name = rs("Holiday_name")
			sEvery_year = rs("Every_year")
%>

	<tr bgcolor="<%=tmpBgColor%>" onmouseover="this.style.background='#FFFCE7'" onmouseout="this.style.background='<%=tmpBgColor%>'">
		<td align="center" width="50"><%=iii%></td>
		<td align="center" width="80"><%=sHoliday%></td>
		<td align="center" width="150"><%=sHoliday_name%></td>
		<td align="center" width="80"><% if sEvery_year="1" then %><img src="/Images/Btn/icon_02.gif" style="cursor:hand;" align="absmiddle"><%else%><img src="/Images/Btn/icon_03.gif" style="cursor:hand;" align="absmiddle"><% end if %></td>
		<td align="center" width="120"><img src="/Images/Btn/BtnEdit.gif" style="cursor:hand;" align="absmiddle" onclick="javascript:parent.fn_popVIEW('<%=sHoliday%>','<%=sHoliday_name%>','<%=sEvery_year%>');">&nbsp;<img src="/Images/Btn/BtnDel.gif" style="cursor:hand;" align="absmiddle" onclick="javascript:fn_del('<%=sHoliday%>',document.inUpFrm);"></td>
		<% if iii = 1 then %>
		<td align="center" rowspan=<%=totalCount%>>매년적용이란 매년 같은 일자의 휴일을 의미함.<br>일자는 수정할 수 없으며, 삭제후 재등록하십시오.</td>
		<% end if %>
	</tr>
<%

			startRow = startRow - 1
			rs.MoveNext
		Loop

		rs.close
		set rs = Nothing

	end if
%>
</table>


<script language="javascript">
<!--
	function ClickBG(f,c,_backcolor){
		for(var i=1; i<=c; i++){
			document.getElementById('cTR' +i).style.backgroundColor = (i==parseInt(f)) ? "#FFEEF9" : "#FFFFFF";
		}
	}
	function fn_del(sHoliday,form)
	{
		if (confirm("선택한 자료를 삭제하시겠습니까?"))
		{
			form.Holiday.value = sHoliday;
			form.guboon.value = "DEL";
			form.submit();
		}
	}
-->
</script>
<!-- #include virtual="/Include/Bottom.asp" -->
