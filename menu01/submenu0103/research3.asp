<!-- #include virtual="/Include/Top.asp" -->
<!-- #include virtual="/Include/PopLayer.asp" -->


<%
	QueryYN = request("QueryYN")
	FromDate = request("FromDate")
	ToDate = request("ToDate")
	whereCD3 = Trim(request("whereCD3"))
	whereCD7 = Trim(request("whereCD7"))

	If QueryYN = "" Then
		whereCD3 = "1"
	End if




%>
<div id='CalendarLayer' style='display:none;'><iframe name='CalendarFrame' src='/include/Calendar.asp' width='172' height='177' border='0' frameborder='0' scrolling='no'></iframe></div>


<table border="0" width="940" cellpadding="0" cellspacing="3" bgcolor="#EFECE5" align="center">
	<tr bgcolor="#FFFFFF">
		<td>
		
			<form method="post" name="inUpFrm" action="<%=Menu_2nd%>" onsubmit="return fn_Search(this);"  style="margin:0">
			<input type="hidden" name="QueryYN" value="">
			<table width="100%" border="0" cellspacing="1" cellpadding="1" align="center" bgcolor="#CCCCCC">
			    <tr>
			        <td width="80" bgcolor="#EFEFEF" class="TDCont">조회기간 :</td>
			        <td colspan="3" bgcolor="#FFFFFF" >
			        	<input value="<%=FromDate%>" name="FromDate" readonly type="text" size="10" onfocus="setFocusColor(this);" onClick="new CalendarFrame.Calendar(this);" onblur="setOutColor(this);">
				    	~
				    	<input value="<%=ToDate%>" name="ToDate" readonly type="text" size="10" onfocus="setFocusColor(this);" onClick="new CalendarFrame.Calendar(this);" onblur="setOutColor(this);">	
			        </td>

			        <td width="80" bgcolor="#EFEFEF" class="TDCont">소속 :</td>
					<td bgcolor="#FFFFFF" nowrap><input type="text" name="sCTIID" readonly value="" maxlength="30" size="30" onfocus="setFocusColor(this);" onblur="setOutColor(this);" align="absmiddle"> <img src="/Images/Comm/IconTip.gif" style="cursor:hand;" align="absmiddle" onClick="pCateSelect('1');" >
					</td>


			        <td colspan='2' rowspan="3" bgcolor="#FFFFFF" align="center">
			        	<img src="/Images/Btn/BtnSearch.gif" style="cursor:hand;" onClick="fn_Search();">
			        	<br><br><img src="/Images/Btn/BtnExcel.gif" style="cursor:hand;" onClick="fn_Xls();">
			        </td>

			    </tr>
				<tr>
			        <td width="80" bgcolor="#EFEFEF" class="TDCont">사건번호 :</td>
			        <td bgcolor="#FFFFFF">
			        	<input value="" name="사건번호" type="text" size="14" onfocus="setFocusColor(this);" onblur="setOutColor(this);"></td>
			        <td width="80" bgcolor="#EFEFEF" class="TDCont">피의자명 :</td>
			        <td bgcolor="#FFFFFF">
			        	<input value="" name="피의자명" type="text" size="14" onfocus="setFocusColor(this);" onblur="setOutColor(this);"></td>

			        <td width="80" bgcolor="#EFEFEF" class="TDCont">담당수사관 :</td>
			        <td bgcolor="#FFFFFF">

			        	<select name="whereCD1" size="1" class="ComboFFFCE7">
							<option value="">선택-----------</option>
							<option value="손민경">손민경</option>
						</select>
					</td>
				</tr
			</table>
			</form>
		</td>
	</tr>
</table>

<table width="940" height="5" border="0" cellspacing="0" cellpadding="0" align="center"><tr><td></td></tr></table>
<table width="940" border="0" cellspacing="1" cellpadding="1" align="center" bgcolor="#CCCCCC">
	<tr height="20" bgcolor="#EEF6FF" align="center">
		<td rowspan=2>No</td>
		<td rowspan=2>사건번호</td>
		<td rowspan=2>사건명</td>
		<td rowspan=2>첨부물<br>유무</td>
		<td rowspan=2>출처일자</td>
		<td colspan=3>담당수사관</td>
		<td colspan=3>연락처유무</td>
		<td rowspan=2>평가<br>점수</td>
		<td rowspan=2>관리</td>

	</tr>
	<tr height="20" bgcolor="#EEF6FF" align="center">
		<td >소속</td>
		<td >계급</td>
		<td >성명</td>
		<td >피의자</td>
		<td >피해자</td>
		<td >지휘관</td>
	</tr>
	<tr><td colspan="17" height="1" bgcolor="#FFFFFF"></td></tr>

	<tr id="cTR1" style="cursor:hand;" bgcolor="#ffffff" onmouseover="this.style.background='#FFFCE7'" bgcolor="#FFFFFF" onmouseout="this.style.background='#FFFFFF'">

			<td align="center">1</td>
			<td align="center">9X09-05-0014</td>
			<td align="center">안전사고-차량-교통사고,도로교통법위반</td>
			<td align="center"><a href="##">사건일지.hwp</a><br><a href="##">사진2.jpg</a></td>
			<td align="center">2009-05-20</td>
			<td align="center">OOO</td>
			<td align="center">병장</td>
			<td align="center">손성훈</td>
			<td align="center"><img src="/Images/Btn/icon_02.gif" style="cursor:hand;" align="absmiddle"></td>
			<td align="center"><img src="/Images/Btn/icon_02.gif" style="cursor:hand;" align="absmiddle"></td>
			<td align="center"><img src="/Images/Btn/icon_02.gif" style="cursor:hand;" align="absmiddle"></td>
			<td align="center">8.15</td>
			<td align="center">
				<img src="/Images/Btn/BtnIconModify.gif" title="수정" style="cursor:hand;" align="absmiddle" onclick="javascript:fn_update('1111','UP');">
				<img src="/Images/Btn/BtnIconDel.gif" title="삭제" style="cursor:hand;" align="absmiddle"onclick="javascript:fn_del('1111','DEL');">
			</td>
	</tr>
	<tr id="cTR1" style="cursor:hand;" bgcolor="#ffffff" onmouseover="this.style.background='#FFFCE7'" bgcolor="#FFFFFF" onmouseout="this.style.background='#FFFFFF'">

			<td align="center">2</td>
			<td align="center">9X09-04-0007</td>
			<td align="center">안전사고-차량-교통사고,도로교통법위반</td>
			<td align="center"></td>
			<td align="center">2009-04-30</td>
			<td align="center">ㅇㅇ군단헌병대</td>
			<td align="center">ㅇㅇ</td>
			<td align="center">ㅇㅇㅇ</td>
			<td align="center"><img src="/Images/Btn/icon_02.gif" style="cursor:hand;" align="absmiddle"></td>
			<td align="center"><img src="/Images/Btn/icon_02.gif" style="cursor:hand;" align="absmiddle"></td>
			<td align="center"><img src="/Images/Btn/icon_02.gif" style="cursor:hand;" align="absmiddle"></td>
			<td align="center">9.0</td>
			<td align="center">
				<img src="/Images/Btn/BtnIconModify.gif" title="수정" style="cursor:hand;" align="absmiddle" onclick="javascript:fn_update('1111','UP');">
				<img src="/Images/Btn/BtnIconDel.gif" title="삭제" style="cursor:hand;" align="absmiddle"onclick="javascript:fn_del('1111','DEL');">
			</td>
	</tr>
	<tr id="cTR1" style="cursor:hand;" bgcolor="#ffffff" onmouseover="this.style.background='#FFFCE7'" bgcolor="#FFFFFF" onmouseout="this.style.background='#FFFFFF'">

			<td align="center">3</td>
			<td align="center">9X09-04-0006</td>
			<td align="center">군폭력</td>
			<td align="center"><a href="##"></a></td>
			<td align="center">2009-04-15</td>
			<td align="center">ㅇㅇ군단헌병대</td>
			<td align="center">ㅇㅇ</td>
			<td align="center">ㅇㅇㅇ</td>
			<td align="center"><img src="/Images/Btn/icon_02.gif" style="cursor:hand;" align="absmiddle"></td>
			<td align="center"><img src="/Images/Btn/icon_02.gif" style="cursor:hand;" align="absmiddle"></td>
			<td align="center"><img src="/Images/Btn/icon_02.gif" style="cursor:hand;" align="absmiddle"></td>
			<td align="center">9.0</td>
			<td align="center">
				<img src="/Images/Btn/BtnIconModify.gif" title="수정" style="cursor:hand;" align="absmiddle" onclick="javascript:fn_update('1111','UP');">
				<img src="/Images/Btn/BtnIconDel.gif" title="삭제" style="cursor:hand;" align="absmiddle"onclick="javascript:fn_del('1111','DEL');">
			</td>
	</tr>

	<tr id="cTR1" style="cursor:hand;" bgcolor="#ffffff" onmouseover="this.style.background='#FFFCE7'" bgcolor="#FFFFFF" onmouseout="this.style.background='#FFFFFF'">

			<td align="center">4</td>
			<td align="center">9X09-04-0001</td>
			<td align="center">단순폭력</td>
			<td align="center"><a href="##"></a></td>
			<td align="center">2009-04-01</td>
			<td align="center">ㅇㅇ군단헌병대</td>
			<td align="center">상병</td>
			<td align="center">김00</td>
			<td align="center"><img src="/Images/Btn/icon_02.gif" style="cursor:hand;" align="absmiddle"></td>
			<td align="center"><img src="/Images/Btn/icon_02.gif" style="cursor:hand;" align="absmiddle"></td>
			<td align="center"><img src="/Images/Btn/icon_02.gif" style="cursor:hand;" align="absmiddle"></td>
			<td align="center">9.0</td>
			<td align="center">
				<img src="/Images/Btn/BtnIconModify.gif" title="수정" style="cursor:hand;" align="absmiddle" onclick="javascript:fn_update('1111','UP');">
				<img src="/Images/Btn/BtnIconDel.gif" title="삭제" style="cursor:hand;" align="absmiddle"onclick="javascript:fn_del('1111','DEL');">
			</td>
	</tr>
</table>

<table border="0" cellpadding="0" cellspacing="0" width="940" align="center">
	<tr><td height="5"></td></tr>
	<tr><td height="1" bgcolor="#D6D6D6"></td></tr>
	<tr height="22" bgcolor="#EEF6FF"><td align="center"></td></tr>
	<tr><td height="1" bgcolor="#D6D6D6"></td></tr>
</table>
<table border="0" cellpadding="0" cellspacing="0" width="940" align="center">
	<tr><td height="5"></td></tr>
</table>

<table border="0" cellpadding="0" cellspacing="0" width="940" align="center">
	<tr><td><iframe frameborder=0 marginheight=0 marginwidth=0 width="100%" height="0" scrolling="no" name="AsInfo1fr"></iframe></td></tr>
</table>



<script>

	function fn_Search() {

		//document.inUpFrm.QueryYN.value = "Y";
		document.inUpFrm.submit();
	}
	
	function fn_Xls() {
		location.href="Part_Xls.asp?<%=pageWHERE%>"
	}

	function pCateSelect(cn){
		Cate1 = '' ; //eval("inUpFrm.ACLASS"+cn).value;
		CUSTNO = '0000000000'; //parent.MemInfoFrame.frmSearch.CUSTNO.value;

		if (cn == '1')
		{//PSEQ1
			Relation = '0';//eval("inUpFrm.RELATION"+cn).value;
			RelationSeq = '0';//eval("inUpFrm.PSEQ"+cn).value;
			GoURL ="/Include/PopUp/PCategory.asp?Cate1=" +Cate1+ "&FM=" +cn+ "&CUSTNO=" +CUSTNO+"&Relation="+Relation+"&RelationSeq="+RelationSeq;
		}
		else
		{
			GoURL ="/Include/PopUp/PCategory.asp?Cate1=" +Cate1+ "&FM=" +cn+ "&CUSTNO=" +CUSTNO;
		}

		ShowPOPLayer(GoURL,'720','470');
	}
</script>

<!-- #include virtual="/Include/Bottom.asp" -->