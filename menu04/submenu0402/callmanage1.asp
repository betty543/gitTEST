<!-- #include virtual="/Include/Top.asp" -->
<%
	sql = "select convert(varchar(19),getdate(),121)"
	set Rs = db.execute(sql)
	sDatetime = rs(0)
%>
<table border="0" width="940" cellpadding="0" cellspacing="3" bgcolor="#EFECE5" align="center">
	<tr bgcolor="#FFFFFF">
		<td>
		
			<form method="post" name="inUpFrm" style="margin:0">
			<input type="hidden" name="QueryYN" value="">
			<table width="100%" border="0" cellspacing="1" cellpadding="1" align="center" bgcolor="#ffffff">
			    <tr>
					<td align="left" bgcolor="#FFFFFF" class="TDCont">◈상담일시: 2009-01-01 15:15</td>
					<td align="right"><img src="/Images/Btn/BtnASRegi.gif" style="cursor:hand;" class="None" align="absmiddle" onClick="fn_inup();"></td>
				</tr>
			</table>
			<table width="100%" border="0" cellspacing="1" cellpadding="1" align="center" bgcolor="#CCCCCC">
			    <tr>
					<td bgcolor="#EEF6FF" class="TDCont" width=150>성별</td>
					<td bgcolor="#FFFFFF" width=200>
						<input type="radio" name="sUSEYN" value="Y" class="none" onClick="fn_YES();" checked > 남
						<input type="radio" name="sUSEYN" value="N" class="none" onClick="fn_YES();" > 녀
					</td>
					<td bgcolor="#EEF6FF" class="TDCont" width=150>상담분야</td>
					<td bgcolor="#FFFFFF" width=200>
						<select name="sGRADE" size="1" class="ComboFFFCE7">
							<Option value ='A등급' selected>성범죄</option>
						</select>					
					</td>
					<td bgcolor="#EEF6FF" class="TDCont" width=150>소속</td>
					<td bgcolor="#FFFFFF" width=200 colspan=3>						<select name="sGRADE" size="1" class="ComboFFFCE7">
							<Option value ='A등급' >소속선택</option>
							<Option value ='A등급' selected>1군</option>
							<Option value ='A등급' >2군</option>
						</select>	
					</td>

				</tr>
			    <tr>
					<td bgcolor="#EEF6FF" class="TDCont" width=150>내담자이름</td>
					<td bgcolor="#FFFFFF" width=200><input type="text" name="sCTIID" value="손민경" maxlength="20" size="20" onfocus="setFocusColor(this);" onblur="setOutColor(this);" align="absmiddle">
					</td>
					<td bgcolor="#EEF6FF" class="TDCont" width=150>계급</td>
					<td bgcolor="#FFFFFF" width=200>
						<select name="sGRADE" size="1" class="ComboFFFCE7">
							<Option value ='A등급' selected>미상</option>
						</select>					
					</td>
					<td bgcolor="#EEF6FF" class="TDCont" width=150>연락처</td>
					<td bgcolor="#FFFFFF" width=200><input type="text" name="sCTIID" value="01051850478" maxlength="20" size="20" onfocus="setFocusColor(this);" onblur="setOutColor(this);" align="absmiddle">
					</td>
					<td bgcolor="#EEF6FF" class="TDCont" width=150>상담방법</td>
					<td bgcolor="#FFFFFF" width=200>						<select name="sGRADE" size="1" class="ComboFFFCE7">
							<Option value ='A등급' selected>군전화</option>
						</select>
					</td>

				</tr>
			    <tr>
					<td bgcolor="#EEF6FF" class="TDCont" width=150>인지경로</td>
					<td bgcolor="#FFFFFF" width=200>						<select name="sGRADE" size="1" class="ComboFFFCE7">
							<Option value ='A등급' selected>인지경로선택</option>
							<Option value ='A등급' selected>군대홈페이지</option>
						</select>
					</td>
					<td bgcolor="#EEF6FF" class="TDCont" width=150>가해자</td>
					<td bgcolor="#FFFFFF" width=200>						<select name="sGRADE" size="1" class="ComboFFFCE7">
							<Option value ='A등급' selected>가해자선택</option>
						</select>
					</td>
					<td bgcolor="#EEF6FF" class="TDCont" width=150>상담관</td>
					<td bgcolor="#FFFFFF" width=200>관리자
					</td>
					<td bgcolor="#EEF6FF" class="TDCont" width=150>조치결과</td>
					<td bgcolor="#FFFFFF" width=200>
						<select name="sGRADE" size="1" class="ComboFFFCE7">
							<Option value ='A등급' selected>조치결과</option>
						</select>					
					</td>
				</tr>

			    <tr>
					<td bgcolor="#EEF6FF" class="TDCont" width=150>상담내용</td>
					<td bgcolor="#FFFFFF" colspan=7 width=850><textarea name="REPLYA1" style="width:100%; height:100" wrap="soft" class="TextareaInput" onKeyUp="fn_SetTEXT('REPLYA','1',this.value); UpdateChar('inUpFrm.REPLYA1',4000,5);"></textarea>			
					</td>
				</tr>
			    <tr>
					<td bgcolor="#EEF6FF" class="TDCont" width=150 >특이사항</td>
					<td bgcolor="#FFFFFF" colspan=7 width=850>	<textarea name="REPLYA1" style="width:100%; height:100" wrap="soft" class="TextareaInput" onKeyUp="fn_SetTEXT('REPLYA','1',this.value); UpdateChar('inUpFrm.REPLYA1',4000,5);"></textarea>			
					</td>
				</tr>
			</table>
			<table width="100%" height="5" border="0" cellspacing="0" cellpadding="0" align="center"><tr><td></td></tr></table>
			</table>
			</form>
		</td>
	</tr>
</table>
<table width="940" height="5" border="0" cellspacing="0" cellpadding="0" align="center"><tr><td></td></tr></table>
<table border="0" width="940" cellspacing="1" cellpadding="0" width="100%" bgcolor="#CCCCCC" align="center">
	<tr height="25" bgcolor="#F3F3F3" align="center">
		<td><b>순번</b></td>
		<td><b>상담일시</b></td>
		<td><b>상담방법</b></td>
		<td><b>소속</b></td>
		<td><b>계급</b></td>
		<td><b>성명</b></td>
		<td><b>상담관</b></td>
		<td><b>성별</b></td>
		<td ><b>상담분야</b></td>
		<td><b>관리</b></td>
	</tr>
	<tr height="25" bgcolor="#ffffff" align="center">
		<td align="center" colspan=10>기존 상담이력이 존재하지 않습니다</td>
	</tr>

		<!--<tr bgcolor="#FFFFFF">
			<td class="TDCont">1</td>
			<td class="TDCont">2009-01-01 15:00</td>
			<td class="TDCont">전화</td>
			<td align="center" width=400>ㅇㅇ사단</td>
			<td align="center">일병</td>
			<td align="center">손민경</td>
			<td align="center">김상담</td>
			<td align="center">남</td>
			<td align="center">복무부적응</td>
			<td align="center">
				<img src="/Images/Btn/BtnIconModify.gif" title="수정" style="cursor:hand;" align="absmiddle" onclick="javascript:fn_update('1111','UP');">
				<img src="/Images/Btn/BtnIconDel.gif" title="삭제" style="cursor:hand;" align="absmiddle"onclick="javascript:fn_del('1111','DEL');">
			</td>
		</tr>
		<tr bgcolor="#FFFFFF">
			<td class="TDCont">2</td>
			<td class="TDCont">2009-01-01 15:00</td>
			<td class="TDCont">내담</td>
			<td align="center" width=400>ㅇㅇ사단</td>
			<td align="center">일병</td>
			<td align="center">손민경</td>
			<td align="center">김상담</td>
			<td align="center">남</td>
			<td align="center">복무부적응</td>
			<td align="center">
				<img src="/Images/Btn/BtnIconModify.gif" title="수정" style="cursor:hand;" align="absmiddle" onclick="javascript:fn_update('1111','UP');">
				<img src="/Images/Btn/BtnIconDel.gif" title="삭제" style="cursor:hand;" align="absmiddle"onclick="javascript:fn_del('1111','DEL');">
			</td>
		</tr>

		<tr bgcolor="#FFFFFF">
			<td class="TDCont">3</td>
			<td class="TDCont">2009-01-01 15:00</td>
			<td class="TDCont">전화</td>
			<td align="center" width=400>ㅇㅇ사단</td>
			<td align="center">일병</td>
			<td align="center">손민경</td>
			<td align="center">김상담</td>
			<td align="center">남</td>
			<td align="center">복무부적응</td>
			<td align="center">
				<img src="/Images/Btn/BtnIconModify.gif" title="수정" style="cursor:hand;" align="absmiddle" onclick="javascript:fn_update('1111','UP');">
				<img src="/Images/Btn/BtnIconDel.gif" title="삭제" style="cursor:hand;" align="absmiddle"onclick="javascript:fn_del('1111','DEL');">
			</td>
		</tr>
		<tr bgcolor="#FFFFFF">
			<td class="TDCont">...</td>
			<td class="TDCont">2009-02-01 15:00</td>
			<td class="TDCont">전화</td>
			<td align="center" width=400>ㅇㅇ사단</td>
			<td align="center">일병</td>
			<td align="center">손민경</td>
			<td align="center">김상담</td>
			<td align="center">남</td>
			<td align="center">복무부적응</td>
			<td align="center">
				<img src="/Images/Btn/BtnIconModify.gif" title="수정" style="cursor:hand;" align="absmiddle" onclick="javascript:fn_update('1111','UP');">
				<img src="/Images/Btn/BtnIconDel.gif" title="삭제" style="cursor:hand;" align="absmiddle"onclick="javascript:fn_del('1111','DEL');">
			</td>
		</tr>-->

</table>
<!-- #include virtual="/Include/Bottom.asp" -->