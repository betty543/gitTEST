<!-- #include virtual="/Include/Top.asp" -->
<table border="0" width="1200" cellpadding="0" cellspacing="3" bgcolor="#EFECE5" align="center">
	<tr bgcolor="#FFFFFF">
		<td>
		
			<form method="post" name="inUpFrm" style="margin:0">
			<input type="hidden" name="QueryYN" value="">
			<table width="100%" border="0" cellspacing="1" cellpadding="1" align="center" bgcolor="#CCCCCC">

			    <tr>
			        <td width="110" bgcolor="#EFEFEF" class="TDCont">조회기간 :</td>
			        <td  bgcolor="#FFFFFF" colspan=3 width=250>
			        	<input value="2009-01-01" name="FromDate" readonly type="text" size="10" onfocus="setFocusColor(this);" onClick="new CalendarFrame.Calendar(this);" onblur="setOutColor(this);">
				    	~
				    	<input value="2009-03-31" name="ToDate" readonly type="text" size="10" onfocus="setFocusColor(this);" onClick="new CalendarFrame.Calendar(this);" onblur="setOutColor(this);">	
			        </td>

					<td bgcolor="#EFEFEF" class="TDCont" width=150>성별</td>
					<td bgcolor="#FFFFFF" width=200 colspan=2>
						<input type="radio" name="sUSEYN" value="Y" class="none" onClick="fn_YES();" checked > 남
						<input type="radio" name="sUSEYN" value="N" class="none" onClick="fn_YES();" > 녀
					</td>

					<td bgcolor="#EFEFEF" class="TDCont" width=150>상담방법</td>
					<td bgcolor="#FFFFFF" width=200 colspan=1>
						<select name="sGRADE" size="1" class="ComboFFFCE7">
							<Option value ='A등급' selected>군전화</option>
						</select>					
					</td>

			        <td colspan='2' rowspan="2" bgcolor="#FFFFFF" align="center">
			        	<img src="/Images/Btn/BtnSearch.gif" style="cursor:hand;" onClick="goSearch(document.inUpFrm);">
			        	<%IF SS_Login_Secgroup="A" Or SS_Login_Secgroup="B" THEN%><br><br><img src="/Images/Btn/BtnExcel.gif" style="cursor:hand;" onClick="fn_Xls();"><%END IF%>
			        </td>
				</tr>
			    <tr>

					<td bgcolor="#EFEFEF" class="TDCont" width=150>의뢰인</td>
					<td bgcolor="#FFFFFF" width=200>
						<select name="sGRADE" size="1" class="ComboFFFCE7">
							<Option value ='A등급' selected>의뢰인선택</option>
						</select>					
					</td>
					<td bgcolor="#EFEFEF" class="TDCont" width=150>상담분야</td>
					<td bgcolor="#FFFFFF" width=200>
						<select name="sGRADE" size="1" class="ComboFFFCE7">
							<Option value ='A등급' selected>상담분야선택</option>
						</select>					
					</td>
					<td bgcolor="#EFEFEF" class="TDCont" width=150>소속</td>
					<td bgcolor="#FFFFFF" width=250 colspan=2 nowrap><input type="text" name="sCTIID" value="" maxlength="15" size="20" onfocus="setFocusColor(this);" onblur="setOutColor(this);" align="absmiddle"> <img src="/Images/Comm/IconTip.gif" style="cursor:hand;" align="absmiddle">
					</td>
					<td bgcolor="#EFEFEF" class="TDCont" width=150>계급</td>
					<td bgcolor="#FFFFFF" width=200>
						<select name="sGRADE" size="1" class="ComboFFFCE7">
							<Option value ='A등급' selected>계급선택</option>
						</select>					
					</td>
			    </tr>
			</table>
			</form>
		</td>
	</tr>
</table>


<table width="1200" height="5" border="0" cellspacing="0" cellpadding="0" align="center"><tr><td></td></tr></table>
<table border="0" width="1200" cellspacing="1" cellpadding="0" width="100%" bgcolor="#CCCCCC" align="center">
	<tr height="25" bgcolor="#F3F3F3" align="center">
		<td><b>순번</b></td>
		<td><b>상담일시</b></td>
		<td><b>상담방법</b></td>
		<td><b>상담횟수</b></td>
		<td><b>소속</b></td>
		<td><b>계급</b></td>
		<td><b>성명</b></td>
		<td><b>상담관</b></td>
		<td><b>관리</b></td>
	</tr>

		<tr bgcolor="#FFFFFF">
			<td class="TDCont">1</td>
			<td class="TDCont" align="center">2009-01-01 15:00</td>
			<td align="center">전화</td>
			<td align="center">1회</td>
			<td align="center" width=400>ㅇㅇ사단</td>
			<td align="center">일병</td>
			<td align="center">손민경</td>
			<td align="center">김상담</td>
			<td align="center">
				<img src="/Images/Btn/BtnIconModify.gif" title="수정" style="cursor:hand;" align="absmiddle" onclick="javascript:fn_update('1111','UP');">
				<img src="/Images/Btn/BtnIconDel.gif" title="삭제" style="cursor:hand;" align="absmiddle"onclick="javascript:fn_del('1111','DEL');">
				<img src="/Images/Comm/IconWrite.gif" title="인쇄" style="cursor:hand;" align="absmiddle"onclick="javascript:fn_del('1111','DEL');">
			</td>
		</tr>
		<tr bgcolor="#FFFFFF">
			<td class="TDCont">2</td>
			<td class="TDCont" align="center">2009-01-01 19:00</td>
			<td align="center">전화</td>
			<td align="center">1회</td>
			<td align="center" width=400>1소속-2소속-3소속</td>
			<td align="center">일병</td>
			<td align="center">손민경</td>
			<td align="center">김상담</td>
			<td align="center">
				<img src="/Images/Btn/BtnIconModify.gif" title="수정" style="cursor:hand;" align="absmiddle" onclick="javascript:fn_update('1111','UP');">
				<img src="/Images/Btn/BtnIconDel.gif" title="삭제" style="cursor:hand;" align="absmiddle"onclick="javascript:fn_del('1111','DEL');">
				<img src="/Images/Comm/IconWrite.gif" title="인쇄" style="cursor:hand;" align="absmiddle"onclick="javascript:fn_del('1111','DEL');">
				
			</td>
		</tr>

		<tr bgcolor="#FFFFFF">
			<td class="TDCont">3</td>
			<td class="TDCont" align="center">2009-01-02 15:00</td>
			<td align="center">전화</td>
			<td align="center">1회</td>
			<td align="center" width=400>1소속-2소속-3소속</td>
			<td align="center">일병</td>
			<td align="center">손민경</td>
			<td align="center">김상담</td>
			<td align="center">
				<img src="/Images/Btn/BtnIconModify.gif" title="수정" style="cursor:hand;" align="absmiddle" onclick="javascript:fn_update('1111','UP');">
				<img src="/Images/Btn/BtnIconDel.gif" title="삭제" style="cursor:hand;" align="absmiddle"onclick="javascript:fn_del('1111','DEL');">
				<img src="/Images/Comm/IconWrite.gif" title="인쇄" style="cursor:hand;" align="absmiddle"onclick="javascript:fn_del('1111','DEL');">
			</td>
		</tr>

		<tr bgcolor="#FFFFFF">
			<td class="TDCont">...</td>
			<td class="TDCont" align="center">2009-01-09 15:00</td>
			<td align="center">전화</td>
			<td align="center">1회</td>
			<td align="center" width=400>1소속-2소속-3소속</td>
			<td align="center">일병</td>
			<td align="center">손민경</td>
			<td align="center">김상담</td>
			<td align="center">
				<img src="/Images/Btn/BtnIconModify.gif" title="수정" style="cursor:hand;" align="absmiddle" onclick="javascript:fn_update('1111','UP');">
				<img src="/Images/Btn/BtnIconDel.gif" title="삭제" style="cursor:hand;" align="absmiddle"onclick="javascript:fn_del('1111','DEL');">
				<img src="/Images/Comm/IconWrite.gif" title="인쇄" style="cursor:hand;" align="absmiddle"onclick="javascript:fn_del('1111','DEL');">
			</td>
		</tr>

</table>
<table width="1200" cellpadding="0" cellspacing="0" width="100%" align="center">
	<tr><td height="2" bgcolor="#f2f2f2"></td></tr>
	<tr><td height="5"></td></tr>
	<tr><td bgcolor="#D6D6D6" height="1"></td></tr>
	<tr><td bgcolor="#F7F7F7" class="TDL10px" height="25">1  2  3  4 </td></tr>
	<tr><td bgcolor="#D6D6D6" height="1"></td></tr>
	<tr>
		<td height="30" class="TDR10px">
			<img src="/Images/Btn/BtnAdd.gif" style="cursor:hand;" align="absmiddle" onclick="javascript:fn_insert();">
		</td>
	</tr>
</table>
<!-- #include virtual="/Include/Bottom.asp" -->