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
					<td align="left" bgcolor="#FFFFFF" class="TDCont">�»���Ͻ�: 2009-01-01 15:15</td>
					<td align="right"><img src="/Images/Btn/BtnASRegi.gif" style="cursor:hand;" class="None" align="absmiddle" onClick="fn_inup();"></td>
				</tr>
			</table>
			<table width="100%" border="0" cellspacing="1" cellpadding="1" align="center" bgcolor="#CCCCCC">
			    <tr>
					<td bgcolor="#EEF6FF" class="TDCont" width=150>����</td>
					<td bgcolor="#FFFFFF" width=200>
						<input type="radio" name="sUSEYN" value="Y" class="none" onClick="fn_YES();" checked > ��
						<input type="radio" name="sUSEYN" value="N" class="none" onClick="fn_YES();" > ��
					</td>
					<td bgcolor="#EEF6FF" class="TDCont" width=150>���о�</td>
					<td bgcolor="#FFFFFF" width=200>
						<select name="sGRADE" size="1" class="ComboFFFCE7">
							<Option value ='A���' selected>������</option>
						</select>					
					</td>
					<td bgcolor="#EEF6FF" class="TDCont" width=150>�Ҽ�</td>
					<td bgcolor="#FFFFFF" width=200 colspan=3>						<select name="sGRADE" size="1" class="ComboFFFCE7">
							<Option value ='A���' >�ҼӼ���</option>
							<Option value ='A���' selected>1��</option>
							<Option value ='A���' >2��</option>
						</select>	
					</td>

				</tr>
			    <tr>
					<td bgcolor="#EEF6FF" class="TDCont" width=150>�������̸�</td>
					<td bgcolor="#FFFFFF" width=200><input type="text" name="sCTIID" value="�չΰ�" maxlength="20" size="20" onfocus="setFocusColor(this);" onblur="setOutColor(this);" align="absmiddle">
					</td>
					<td bgcolor="#EEF6FF" class="TDCont" width=150>���</td>
					<td bgcolor="#FFFFFF" width=200>
						<select name="sGRADE" size="1" class="ComboFFFCE7">
							<Option value ='A���' selected>�̻�</option>
						</select>					
					</td>
					<td bgcolor="#EEF6FF" class="TDCont" width=150>����ó</td>
					<td bgcolor="#FFFFFF" width=200><input type="text" name="sCTIID" value="01051850478" maxlength="20" size="20" onfocus="setFocusColor(this);" onblur="setOutColor(this);" align="absmiddle">
					</td>
					<td bgcolor="#EEF6FF" class="TDCont" width=150>�����</td>
					<td bgcolor="#FFFFFF" width=200>						<select name="sGRADE" size="1" class="ComboFFFCE7">
							<Option value ='A���' selected>����ȭ</option>
						</select>
					</td>

				</tr>
			    <tr>
					<td bgcolor="#EEF6FF" class="TDCont" width=150>�������</td>
					<td bgcolor="#FFFFFF" width=200>						<select name="sGRADE" size="1" class="ComboFFFCE7">
							<Option value ='A���' selected>������μ���</option>
							<Option value ='A���' selected>����Ȩ������</option>
						</select>
					</td>
					<td bgcolor="#EEF6FF" class="TDCont" width=150>������</td>
					<td bgcolor="#FFFFFF" width=200>						<select name="sGRADE" size="1" class="ComboFFFCE7">
							<Option value ='A���' selected>�����ڼ���</option>
						</select>
					</td>
					<td bgcolor="#EEF6FF" class="TDCont" width=150>����</td>
					<td bgcolor="#FFFFFF" width=200>������
					</td>
					<td bgcolor="#EEF6FF" class="TDCont" width=150>��ġ���</td>
					<td bgcolor="#FFFFFF" width=200>
						<select name="sGRADE" size="1" class="ComboFFFCE7">
							<Option value ='A���' selected>��ġ���</option>
						</select>					
					</td>
				</tr>

			    <tr>
					<td bgcolor="#EEF6FF" class="TDCont" width=150>��㳻��</td>
					<td bgcolor="#FFFFFF" colspan=7 width=850><textarea name="REPLYA1" style="width:100%; height:100" wrap="soft" class="TextareaInput" onKeyUp="fn_SetTEXT('REPLYA','1',this.value); UpdateChar('inUpFrm.REPLYA1',4000,5);"></textarea>			
					</td>
				</tr>
			    <tr>
					<td bgcolor="#EEF6FF" class="TDCont" width=150 >Ư�̻���</td>
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
		<td><b>����</b></td>
		<td><b>����Ͻ�</b></td>
		<td><b>�����</b></td>
		<td><b>�Ҽ�</b></td>
		<td><b>���</b></td>
		<td><b>����</b></td>
		<td><b>����</b></td>
		<td><b>����</b></td>
		<td ><b>���о�</b></td>
		<td><b>����</b></td>
	</tr>
	<tr height="25" bgcolor="#ffffff" align="center">
		<td align="center" colspan=10>���� ����̷��� �������� �ʽ��ϴ�</td>
	</tr>

		<!--<tr bgcolor="#FFFFFF">
			<td class="TDCont">1</td>
			<td class="TDCont">2009-01-01 15:00</td>
			<td class="TDCont">��ȭ</td>
			<td align="center" width=400>�������</td>
			<td align="center">�Ϻ�</td>
			<td align="center">�չΰ�</td>
			<td align="center">����</td>
			<td align="center">��</td>
			<td align="center">����������</td>
			<td align="center">
				<img src="/Images/Btn/BtnIconModify.gif" title="����" style="cursor:hand;" align="absmiddle" onclick="javascript:fn_update('1111','UP');">
				<img src="/Images/Btn/BtnIconDel.gif" title="����" style="cursor:hand;" align="absmiddle"onclick="javascript:fn_del('1111','DEL');">
			</td>
		</tr>
		<tr bgcolor="#FFFFFF">
			<td class="TDCont">2</td>
			<td class="TDCont">2009-01-01 15:00</td>
			<td class="TDCont">����</td>
			<td align="center" width=400>�������</td>
			<td align="center">�Ϻ�</td>
			<td align="center">�չΰ�</td>
			<td align="center">����</td>
			<td align="center">��</td>
			<td align="center">����������</td>
			<td align="center">
				<img src="/Images/Btn/BtnIconModify.gif" title="����" style="cursor:hand;" align="absmiddle" onclick="javascript:fn_update('1111','UP');">
				<img src="/Images/Btn/BtnIconDel.gif" title="����" style="cursor:hand;" align="absmiddle"onclick="javascript:fn_del('1111','DEL');">
			</td>
		</tr>

		<tr bgcolor="#FFFFFF">
			<td class="TDCont">3</td>
			<td class="TDCont">2009-01-01 15:00</td>
			<td class="TDCont">��ȭ</td>
			<td align="center" width=400>�������</td>
			<td align="center">�Ϻ�</td>
			<td align="center">�չΰ�</td>
			<td align="center">����</td>
			<td align="center">��</td>
			<td align="center">����������</td>
			<td align="center">
				<img src="/Images/Btn/BtnIconModify.gif" title="����" style="cursor:hand;" align="absmiddle" onclick="javascript:fn_update('1111','UP');">
				<img src="/Images/Btn/BtnIconDel.gif" title="����" style="cursor:hand;" align="absmiddle"onclick="javascript:fn_del('1111','DEL');">
			</td>
		</tr>
		<tr bgcolor="#FFFFFF">
			<td class="TDCont">...</td>
			<td class="TDCont">2009-02-01 15:00</td>
			<td class="TDCont">��ȭ</td>
			<td align="center" width=400>�������</td>
			<td align="center">�Ϻ�</td>
			<td align="center">�չΰ�</td>
			<td align="center">����</td>
			<td align="center">��</td>
			<td align="center">����������</td>
			<td align="center">
				<img src="/Images/Btn/BtnIconModify.gif" title="����" style="cursor:hand;" align="absmiddle" onclick="javascript:fn_update('1111','UP');">
				<img src="/Images/Btn/BtnIconDel.gif" title="����" style="cursor:hand;" align="absmiddle"onclick="javascript:fn_del('1111','DEL');">
			</td>
		</tr>-->

</table>
<!-- #include virtual="/Include/Bottom.asp" -->