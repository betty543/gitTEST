
<!-- #include virtual="/Include/Top.asp" -->
<table border="0" width="1200" cellpadding="0" cellspacing="3" bgcolor="#EFECE5" align="center">
	<tr bgcolor="#FFFFFF">
		<td>
		
			<form method="post" name="inUpFrm" style="margin:0">
			<input type="hidden" name="QueryYN" value="">
			<table width="100%" border="0" cellspacing="1" cellpadding="1" align="center" bgcolor="#CCCCCC">

			    <tr>
			        <td width="110" bgcolor="#EEF6FF" class="TDCont">��ȸ�Ⱓ :</td>
			        <td  bgcolor="#FFFFFF" colspan=3 width=200>
			        	<input value="2009-01-01" name="FromDate" readonly type="text" size="10" onfocus="setFocusColor(this);" onClick="new CalendarFrame.Calendar(this);" onblur="setOutColor(this);">
				    	~
				    	<input value="2009-03-31" name="ToDate" readonly type="text" size="10" onfocus="setFocusColor(this);" onClick="new CalendarFrame.Calendar(this);" onblur="setOutColor(this);">	
			        </td>

						<td width="110" bgcolor="#EEF6FF" class="TDCont">�߼���</td>
						<td bgcolor="#FFFFFF">
							<select name="sGRADE" size="1" class="ComboFFFCE7">
								<Option value ='A���' selected>�߼��� ����----</option>
							</select>					
						</td>

						<td width="110" bgcolor="#EEF6FF" class="TDCont">�޴�����ȣ</td>
						<td bgcolor="#FFFFFF"><input type="text" name="sCTIID" value="" maxlength="20" size="20" onfocus="setFocusColor(this);" onblur="setOutColor(this);" align="absmiddle"></td>


						<td width="110" bgcolor="#EEF6FF" class="TDCont">������</td>
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


<table width="1200" height="5" border="0" cellspacing="0" cellpadding="0" align="center"><tr><td></td></tr></table>
<table border="0" width="1200" cellspacing="1" cellpadding="0" width="100%" bgcolor="#CCCCCC" align="center">
	<tr height="25" bgcolor="#F3F3F3" align="center">
		<td><b>����</b></td>
		<td><b>���ۿ�û�Ͻ�</b></td>
		<td><b>��ǹ�ȣ</b></td>
		<td><b>��Ǹ�</b></td>
		<td><b>�߼���</b></td>
		<td><b>�����޴���</b></td>
		<td><b>������</b></td>
		<td><b>�߼۳���</b></td>
		<td ><b>�߼۰��</b></td>
		<td><b>�߼��Ͻ�</b></td>
		<td><b>����</b></td>
	</tr>

		<tr bgcolor="#FFFFFF">
			<td class="TDCont" align="center">1</td>
			<td class="TDCont" align="center">2009-01-01 15:00</td>
			<td align="center">0000000000</td>
			<td align="center"></td>
			<td align="center">�չΰ�</td>
			<td align="center">01051850478</td>
			<td align="center">�չΰ�</td>
			<td align="center">sms�׽�Ʈ</td>
			<td align="center">����</td>
			<td align="center">2009-01-01 15:00</td>
			<td align="center">
				<img src="/Images/Btn/BtnIconModify.gif" title="����" style="cursor:hand;" align="absmiddle" onclick="javascript:fn_update('1111','UP');">
				<img src="/Images/Btn/BtnIconDel.gif" title="����" style="cursor:hand;" align="absmiddle"onclick="javascript:fn_del('1111','DEL');">
			</td>
		
		</tr>
		<tr bgcolor="#FFFFFF">
			<td class="TDCont" align="center">2</td>
			<td class="TDCont" align="center">2009-01-01 15:00</td>
			<td align="center">0000000000</td>
			<td align="center"></td>
			<td align="center">�չΰ�</td>
			<td align="center">01051850478</td>
			<td align="center">�չΰ�</td>
			<td align="center">sms�׽�Ʈ</td>
			<td align="center">����</td>
			<td align="center">2009-01-01 15:00</td>
			<td align="center">
				<img src="/Images/Btn/BtnIconModify.gif" title="����" style="cursor:hand;" align="absmiddle" onclick="javascript:fn_update('1111','UP');">
				<img src="/Images/Btn/BtnIconDel.gif" title="����" style="cursor:hand;" align="absmiddle"onclick="javascript:fn_del('1111','DEL');">
			</td>
		
		</tr>
		<tr bgcolor="#FFFFFF">
			<td class="TDCont" align="center">3</td>
			<td class="TDCont" align="center">2009-01-01 15:00</td>
			<td align="center">0000000000</td>
			<td align="center"></td>
			<td align="center">�չΰ�</td>
			<td align="center">01051850478</td>
			<td align="center">�չΰ�</td>
			<td align="center">sms�׽�Ʈ</td>
			<td align="center">����</td>
			<td align="center">2009-01-01 15:00</td>
			<td align="center">
				<img src="/Images/Btn/BtnIconModify.gif" title="����" style="cursor:hand;" align="absmiddle" onclick="javascript:fn_update('1111','UP');">
				<img src="/Images/Btn/BtnIconDel.gif" title="����" style="cursor:hand;" align="absmiddle"onclick="javascript:fn_del('1111','DEL');">
			</td>
		
		</tr>
		<tr bgcolor="#FFFFFF">
			<td class="TDCont" align="center">4</td>
			<td class="TDCont" align="center">2009-01-01 15:00</td>
			<td align="center">0000000000</td>
			<td align="center"></td>
			<td align="center">�չΰ�</td>
			<td align="center">01051850478</td>
			<td align="center">�չΰ�</td>
			<td align="center">sms�׽�Ʈ</td>
			<td align="center">����</td>
			<td align="center">2009-01-01 15:00</td>
			<td align="center">
				<img src="/Images/Btn/BtnIconModify.gif" title="����" style="cursor:hand;" align="absmiddle" onclick="javascript:fn_update('1111','UP');">
				<img src="/Images/Btn/BtnIconDel.gif" title="����" style="cursor:hand;" align="absmiddle"onclick="javascript:fn_del('1111','DEL');">
			</td>
		
		</tr>
		<tr bgcolor="#FFFFFF">
			<td class="TDCont" align="center">5</td>
			<td class="TDCont" align="center">2009-01-01 15:00</td>
			<td align="center">0000000000</td>
			<td align="center"></td>
			<td align="center">�չΰ�</td>
			<td align="center">01051850478</td>
			<td align="center">�չΰ�</td>
			<td align="center">sms�׽�Ʈ</td>
			<td align="center">����</td>
			<td align="center">2009-01-01 15:00</td>
			<td align="center">
				<img src="/Images/Btn/BtnIconModify.gif" title="����" style="cursor:hand;" align="absmiddle" onclick="javascript:fn_update('1111','UP');">
				<img src="/Images/Btn/BtnIconDel.gif" title="����" style="cursor:hand;" align="absmiddle"onclick="javascript:fn_del('1111','DEL');">
			</td>		
		</tr>


</table>
<table width="1200" cellpadding="0" cellspacing="0" width="100%" align="center">
	<tr><td height="2" bgcolor="#f2f2f2"></td></tr>
	<tr><td height="5"></td></tr>
	<tr><td bgcolor="#D6D6D6" height="1"></td></tr>
	<tr><td bgcolor="#F7F7F7" class="TDL10px" height="25" align="center">1  2  3  4 </td></tr>
	<tr><td bgcolor="#D6D6D6" height="1"></td></tr>
	<tr>
		<td height="30" class="TDR10px">
			
		</td>
	</tr>
</table>

<!-- #include virtual="/Include/Bottom.asp" -->