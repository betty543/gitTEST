<!-- #include virtual="/Include/Top.asp" -->
<table border="0" width="1200" cellpadding="0" cellspacing="3" bgcolor="#EFECE5" align="center">
	<tr bgcolor="#FFFFFF">
		<td>
		
			<form method="post" name="inUpFrm" style="margin:0">
			<input type="hidden" name="QueryYN" value="">
			<table width="100%" border="0" cellspacing="1" cellpadding="1" align="center" bgcolor="#CCCCCC">

			    <tr>
						<td bgcolor="#EEF6FF" class="TDCont" width=150>益血</td>
						<td bgcolor="#FFFFFF" colspan=4>
							<select name="sGRADE" size="1" class="ComboFFFCE7">
								<Option value ='A去厭' selected>益血識澱------------</option>
							</select>					
						</td>

						<td bgcolor="#EEF6FF" class="TDCont" width=150>社紗</td>
						<td bgcolor="#FFFFFF" width=200 colspan=2><input type="text" name="sCTIID" value="" maxlength="15" size="20" onfocus="setFocusColor(this);" onblur="setOutColor(this);" align="absmiddle"> <img src="/Images/Comm/IconTip.gif" style="cursor:hand;" align="absmiddle">
						</td>

						<td bgcolor="#EEF6FF" class="TDCont">穿鉢腰硲</td>
						<td bgcolor="#FFFFFF"><input type="text" name="sCTIID" value="" maxlength="20" size="20" onfocus="setFocusColor(this);" onblur="setOutColor(this);" align="absmiddle"></td>


						<td bgcolor="#EEF6FF" class="TDCont">失誤</td>
						<td bgcolor="#FFFFFF"><input type="text" name="sCTIID" value="" maxlength="20" size="20" onfocus="setFocusColor(this);" onblur="setOutColor(this);" align="absmiddle"></td>


			        <td colspan='2' rowspan="2" bgcolor="#FFFFFF" align="center">
			        	<img src="/Images/Btn/BtnSearch.gif" style="cursor:hand;" onClick="goSearch(document.inUpFrm);">
			        	<%IF SS_Login_Secgroup="A" Or SS_Login_Secgroup="B" THEN%><br><br><img src="/Images/Btn/BtnExcel.gif" style="cursor:hand;" onClick="fn_Xls();"><%END IF%>
			        </td>
				</tr>

			</table>
			</form>
		</td>
	</tr>
</table>
<table width="1200" border="0" cellpadding="0" cellspacing="0" align="center">
	<tr valign="top">
		<td width="800" height="750">
			<!--<iframe src="User_List.html" name="ListFrame" width="100%" height="100%" frameborder=0 marginheight=0 marginwidth=0 scrolling="no"></iframe>-->

	
			<table width="100%" border="0" cellspacing="0" cellpadding="0">
        		<tr><td height="22" colspan="2" class="FBlk">�� <b>紫遂切鯉系</b></td></tr>
        	</table>
        	<table width="100%" border="0" cellspacing="1" cellpadding="1" bgcolor="#CCCCCC">
        		<tr height="20" bgcolor="#EEF6FF" align="center">
        			<td>企雌切</td>
        			<td>NO</td>
        			<td>益血</td>
        			<td>社紗</td>
        			<td>域厭</td>
        			<td>浦腰</td>
        			<td>失誤</td>
        			<td>穿鉢腰硲</td>
        			<td></td>
        		</tr>
        		<tr><td colspan="11" height="1" bgcolor="#FFFFFF"></td></tr>

				<tr onmouseover="this.style.background='#FFFCE7'" bgcolor="#FFFFFF" >
					<td align="center"><img src="/Images/Btn/icon_02.gif" style="cursor:hand;" align="absmiddle"></td>
					<td align="center">1</td>
					<td align="center">A益血</td>
					<td align="center">社紗1-社紗2-社紗3</td>
					<td align="center"> </td>
					<td align="center"> </td>
					<td align="center">謝肯井</td>
					<td align="center">010-234-1234</td>
					<td align="center"><img src="/Images/Btn/icon_02.gif" style="cursor:hand;" align="absmiddle"></td>
				</tr>
				<tr onmouseover="this.style.background='#FFFCE7'" bgcolor="#FFFFFF" >
					<td align="center"><img src="/Images/Btn/icon_02.gif" style="cursor:hand;" align="absmiddle"></td>
					<td align="center">2</td>
					<td align="center">A益血</td>
					<td align="center">社紗1-社紗2-社紗3</td>
					<td align="center"> </td>
					<td align="center"> </td>
					<td align="center">沿焼巷鯵</td>
					<td align="center">010-234-1234</td>
					<td align="center"><img src="/Images/Btn/icon_02.gif" style="cursor:hand;" align="absmiddle"></td>
				</tr>


				<tr bgcolor="#FFFFFF" >
					<td align="center" colspan=11 ><img src="/Images/Btn/BtnPlus.gif" style="cursor:hand;" align="absmiddle"></td>
				</tr>

        	</table>       	




		</td>
		<td width="10" height="750" align='center' valign='center'>
			<!--<iframe src="User_List.html" name="ListFrame" width="100%" height="100%" frameborder=0 marginheight=0 marginwidth=0 scrolling="no"></iframe>-->

	
			<!--<table width="80%" border="0" cellspacing="0" cellpadding="0">
        		<tr><td height="22" colspan="2" class="FBlk"></td></tr>
        	</table>
        	<table width="80%" border="0" cellspacing="1" cellpadding="1" bgcolor="#CCCCCC">
        		<tr height="20" bgcolor="#ffffff" align="center">
        			<td></td>        		
        		</tr>
        	</table> -->      	




		</td>
		<td width="400" height="750">
			<!--<iframe src="User_List.html" name="ListFrame" width="100%" height="100%" frameborder=0 marginheight=0 marginwidth=0 scrolling="no"></iframe>-->

	
			<table width="100%" border="0" cellspacing="0" cellpadding="0">
        		<tr><td height="22" colspan="2" class="FBlk">�� <b>搭鉢企雌</b></td></tr>
        	</table>
        	<table width="100%" border="0" cellspacing="1" cellpadding="1" bgcolor="#CCCCCC">
        		<tr height="20" bgcolor="#EEF6FF" align="center">
        			<td>企雌切</td>
        			<td>NO</td>
        			<td>社紗</td>
        			<td>域厭</td>
        			<td>浦腰</td>
        			<td>失誤</td>
        		</tr>
        		<tr><td colspan="6" height="1" bgcolor="#FFFFFF"></td></tr>

				<tr onmouseover="this.style.background='#FFFCE7'" bgcolor="#FFFFFF" >
					<td align="center"><img src="/Images/Btn/icon_02.gif" style="cursor:hand;" align="absmiddle"></td>
					<td align="center">1</td>
					<td align="center">社紗1-社紗2-社紗3</td>
					<td align="center"> </td>
					<td align="center"> </td>
					<td align="center">謝肯井</td>
				</tr>
				<tr onmouseover="this.style.background='#FFFCE7'" bgcolor="#FFFFFF" >
					<td align="center"><img src="/Images/Btn/icon_02.gif" style="cursor:hand;" align="absmiddle"></td>
					<td align="center">2</td>
					<td align="center">社紗1-社紗2-社紗3</td>
					<td align="center"> </td>
					<td align="center"> </td>
					<td align="center">沿焼巷鯵</td>
				</tr>

				<tr onmouseover="this.style.background='#FFFCE7'" bgcolor="#FFFFFF" >
					<td align="center"><img src="/Images/Btn/icon_02.gif" style="cursor:hand;" align="absmiddle"></td>
					<td align="center">3</td>
					<td align="center">社紗1-社紗2-社紗3</td>
					<td align="center"> </td>
					<td align="center"> </td>
					<td align="center">ししし</td>
				</tr>
				<tr onmouseover="this.style.background='#FFFCE7'" bgcolor="#FFFFFF" >
					<td align="center"><img src="/Images/Btn/icon_02.gif" style="cursor:hand;" align="absmiddle"></td>
					<td align="center">4</td>
					<td align="center">社紗1-社紗2-社紗3</td>
					<td align="center"> </td>
					<td align="center"> </td>
					<td align="center">けけけ</td>
				</tr>

				<tr onmouseover="this.style.background='#FFFCE7'" bgcolor="#FFFFFF" >
					<td align="center"><img src="/Images/Btn/icon_02.gif" style="cursor:hand;" align="absmiddle"></td>
					<td align="center">5</td>
					<td align="center">社紗1-社紗2-社紗3</td>
					<td align="center"> </td>
					<td align="center"> </td>
					<td align="center">けけけ</td>
				</tr>

				<tr onmouseover="this.style.background='#FFFCE7'" bgcolor="#FFFFFF" >
					<td align="center"><img src="/Images/Btn/icon_02.gif" style="cursor:hand;" align="absmiddle"></td>
					<td align="center">...</td>
					<td align="center">社紗1-社紗2-社紗3</td>
					<td align="center"> </td>
					<td align="center"> </td>
					<td align="center">沿姶紫</td>
				</tr>

				<tr onmouseover="this.style.background='#FFFCE7'" bgcolor="#FFFFFF" >
					<td align="center"><img src="/Images/Btn/icon_02.gif" style="cursor:hand;" align="absmiddle"></td>
					<td align="center">30</td>
					<td align="center">社紗1-社紗2-社紗3</td>
					<td align="center"> </td>
					<td align="center"> </td>
					<td align="center">戚雌紫</td>
				</tr>

				<tr bgcolor="#FFFFFF" >
					<td align="center" colspan=6><img src="/Images/Btn/BtnMinus.gif" style="cursor:hand;" align="absmiddle"> <img src="/Images/Btn/BtnCallSend.gif" style="cursor:hand;" align="absmiddle"></td>
				</tr>

        	</table>       	


		</td>

	</tr>
</table>
<!-- #include virtual="/Include/Bottom.asp" -->