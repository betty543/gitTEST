
<!-- #include virtual="/Include/Top.asp" -->
<table width="940" border="1" cellpadding="0" cellspacing="0" align="center">
	<tr valign="top">

		<td width="340">
			<table width="100%" border="0" cellspacing="0" cellpadding="0">
        		<tr><td height="22" colspan="2" class="FBlk">◈ <b>그룹</b></td><td colspan="5" align="right" height=28><img src="/Images/Btn/BtnAdd.gif" title="그룹추가" style="cursor:hand;" align="absmiddle" onclick="javascript:fn_update('1111','UP');">&nbsp;</td></tr>
        	</table>
        	<table width="100%" border="0" cellspacing="1" cellpadding="1" bgcolor="#CCCCCC">
        		<tr height="20" bgcolor="#EEF6FF" align="center">
        			<td>NO</td>
        			<td>구분</td>
        			<td>그룹명</td>
        			<td>사용여부</td>
        			<td width=40 align='center'>관리</td>
        		</tr>
        		<tr><td colspan="5" height="1" bgcolor="#FFFFFF"></td></tr>
				<tr onmouseover="this.style.background='#FFFCE7'" bgcolor="#FFFFFF" onmouseout="this.style.background='#FFFFFF'">
					<td align="center">1</td>
					<td align="center">개인</td>
					<td align="center">친구</td>
					<td align="center">사용</td>

					<td align="center">
							<img src="/Images/Btn/BtnIconModify.gif" title="그룹수정" style="cursor:hand;" align="absmiddle" onclick="javascript:fn_update('1111','UP');">
							<img src="/Images/Btn/BtnIconDel.gif" title="그룹삭제" style="cursor:hand;" align="absmiddle" onclick="javascript:fn_update('1111','DEL');">
					</td>
				</tr>
				<tr onmouseover="this.style.background='#FFFCE7'" bgcolor="#FFFFFF" onmouseout="this.style.background='#FFFFFF'">
					<td align="center">2</td>
					<td align="center">공통</td>
					<td align="center">긴급발령</td>
					<td align="center">사용</td>

					<td align="center">
							<img src="/Images/Btn/BtnIconModify.gif" title="그룹수정" style="cursor:hand;" align="absmiddle" onclick="javascript:fn_update('1111','UP');">
							<img src="/Images/Btn/BtnIconDel.gif" title="그룹삭제" style="cursor:hand;" align="absmiddle" onclick="javascript:fn_update('1111','DEL');">
					</td>
				</tr>
			</table>

		</td>
		<td width="10"></td>
		<td width="590">
			<!--<iframe src="User_List.html" name="ListFrame" width="100%" height="100%" frameborder=0 marginheight=0 marginwidth=0 scrolling="no"></iframe>-->

	
			<table width="100%" border="0" cellspacing="0" cellpadding="0">
        		<tr><td height="22" colspan="2" class="FBlk">◈ <b>(선택한그룹의)주소록</b></td><td height=28 colspan="1" height="1" align="right"><img src="/Images/Btn/BtnAdd.gif" title="주소록추가" style="cursor:hand;" align="absmiddle" onclick="javascript:fn_update('1111','UP');">&nbsp;</td></tr>
        	</table>
        	<table width="100%" border="0" cellspacing="1" cellpadding="1" bgcolor="#CCCCCC">
        		<tr height="20" bgcolor="#EEF6FF" align="center">
        			<td>선택</td>
        			<td>NO</td>
        			<td>소속</td>
        			<td>계급</td>
        			<td>성명</td>
        			<td>휴대폰번호</td>
        			<td width=40 align='center'>관리</td>

        		</tr>
        		<tr><td colspan="7" height="1" bgcolor="#FFFFFF"></td></tr>

				<tr onmouseover="this.style.background='#FFFCE7'" bgcolor="#FFFFFF" onmouseout="this.style.background='#FFFFFF'" >
					<td align="center"><img src="/Images/Btn/icon_02.gif" style="cursor:hand;" align="absmiddle" ></td>
					<td align="center">1</td>
					<td align="center" width=200>소속1-소속2-소속3</td>
					<td align="center"> </td>
					<td align="center">손민경</td>
					<td align="center">01051850478</td>
					<td align="center">
						<img src="/Images/Btn/BtnIconModify.gif" title="수정" style="cursor:hand;" align="absmiddle" onclick="javascript:fn_update('1111','UP');">
						<img src="/Images/Btn/BtnIconDel.gif" title="수정" style="cursor:hand;" align="absmiddle" onclick="javascript:fn_update('1111','UP');">
					</td>
				</tr>
				<tr onmouseover="this.style.background='#FFFCE7'" bgcolor="#FFFFFF" onmouseover="this.style.background='#FFFCE7'" onmouseout="this.style.background='#FFFFFF'" >
					<td align="center"><img src="/Images/Btn/icon_02.gif" style="cursor:hand;" align="absmiddle"></td>
					<td align="center">2</td>
					<td align="center">소속1-소속2-소속3</td>
					<td align="center"> </td>
					<td align="center">OOO</td>
					<td align="center">00000000000</td>
					<td align="center">
						<img src="/Images/Btn/BtnIconModify.gif" title="수정" style="cursor:hand;" align="absmiddle" onclick="javascript:fn_update('1111','UP');">
						<img src="/Images/Btn/BtnIconDel.gif" title="수정" style="cursor:hand;" align="absmiddle" onclick="javascript:fn_update('1111','UP');">
					</td>
				</tr>
        	</table>    
			<table width="100%" border="0" cellspacing="0" cellpadding="0">
        		<tr><td height="22" colspan="2" class="FBlk"></b></td></tr>
        	</table>			
        	<table width="100%" border="0" cellspacing="1" cellpadding="1" bgcolor="#CCCCCC">
        		<tr bgcolor="#EEF6FF" height="40"><td height="22" colspan="2" class="FBlk" align="center"> 선택한 대상을
							<select name="whereCD1" size="1" class="ComboFFFCE7">
								<option value="">개인-친구</option>
							</select>에 
							<input type="button" name="BtnPlay" value="추가하기" style="width:120; height:20%;" class="Btn3" onClick="javascript:fn_Player();">  또는 <input type="button" name="BtnPlay" value="이동하기" style="width:120; height:20%;" class="Btn3" onClick="javascript:fn_Player();"></b></td></tr>
        	</table>


		</td>

	</tr>
</table>
