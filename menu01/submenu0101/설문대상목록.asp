
<!-- #include virtual="/Include/Top.asp" -->
<table border="0" width="1200" cellpadding="0" cellspacing="3" bgcolor="#EFECE5" align="center">
	<tr bgcolor="#FFFFFF">
		<td>
		
			<form method="post" name="inUpFrm" style="margin:0">
			<input type="hidden" name="QueryYN" value="">
			<table width="100%" border="0" cellspacing="1" cellpadding="1" align="center" bgcolor="#CCCCCC">
			    <tr>
			        <td width="80" bgcolor="#EFEFEF" class="TDCont">��ȸ�Ⱓ :</td>
			        <td colspan="3" bgcolor="#FFFFFF" >
			        	<input value="2009-01-01" name="FromDate" readonly type="text" size="10" onfocus="setFocusColor(this);" onClick="new CalendarFrame.Calendar(this);" onblur="setOutColor(this);">
				    	~
				    	<input value="2009-03-31" name="ToDate" readonly type="text" size="10" onfocus="setFocusColor(this);" onClick="new CalendarFrame.Calendar(this);" onblur="setOutColor(this);">	
			        </td>
			        <td width="80" bgcolor="#EFEFEF" class="TDCont">��ǹ�ȣ :</td>
			        <td bgcolor="#FFFFFF">
			        	<input value="" name="FromDate" readonly type="text" size="10" onfocus="setFocusColor(this);" onClick="new CalendarFrame.Calendar(this);" onblur="setOutColor(this);"></td>
			        <td width="80" bgcolor="#EFEFEF" class="TDCont">�Ҽ� :</td>
					<td bgcolor="#FFFFFF" nowrap><input type="text" name="sCTIID" value="" maxlength="15" size="20" onfocus="setFocusColor(this);" onblur="setOutColor(this);" align="absmiddle"> <img src="/Images/Comm/IconTip.gif" style="cursor:hand;" align="absmiddle">
					</td>
			        <td width="80" bgcolor="#EFEFEF" class="TDCont">�����ڸ� :</td>
			        <td bgcolor="#FFFFFF">
			        	<input value="" name="FromDate" readonly type="text" size="10" onfocus="setFocusColor(this);" onClick="new CalendarFrame.Calendar(this);" onblur="setOutColor(this);"></td>

			        <td width="80" bgcolor="#EFEFEF" class="TDCont">������� :</td>
			        <td bgcolor="#FFFFFF">

			        	<select name="whereCD1" size="1" class="ComboFFFCE7">
							<option value="">����</option>
						</select>
					</td>

			        <td colspan='2' rowspan="3" bgcolor="#FFFFFF" align="center">
			        	<img src="/Images/Btn/BtnSearch.gif" style="cursor:hand;" onClick="goSearch(document.inUpFrm);">
			        	<br><br><img src="/Images/Btn/BtnExcel.gif" style="cursor:hand;" onClick="fn_Xls();">
			        </td>

			    </tr>
			</table>
			</form>
		</td>
	</tr>
</table>

<table width="1200" height="5" border="0" cellspacing="0" cellpadding="0" align="center"><tr><td></td></tr></table>

<table width="1200" border="0" cellspacing="1" cellpadding="1" align="center" bgcolor="#CCCCCC">
	<tr height="20" bgcolor="#EEF6FF" align="center">
		<td rowspan=2>No</td>
		<td rowspan=2>��ǹ�ȣ</td>
		<td rowspan=2>��Ǹ�</td>
		<td rowspan=2>÷�ι�<br>����</td>
		<td rowspan=2>��ó����</td>
		<td colspan=3>�������</td>
		<td colspan=3>����ó����</td>
		<td rowspan=2>����</td>

	</tr>
	<tr height="20" bgcolor="#EEF6FF" align="center">
		<td >�Ҽ�</td>
		<td >���</td>
		<td >����</td>
		<td >������</td>
		<td >������</td>
		<td >���ְ�</td>
	</tr>
	<tr><td colspan="17" height="1" bgcolor="#FFFFFF"></td></tr>


	<tr id="cTR1" style="cursor:hand;" bgcolor="#ffffff" >

			<td align="center">1</td>
			<td align="center">0000000000</td>
			<td align="center">������-����</td>
			<td align="center"><img src="/Images/Btn/icon_02.gif" style="cursor:hand;" align="absmiddle"></td>
			<td align="center">2009-01-01</td>
			<td align="center">���������庴��</td>
			<td align="center">����</td>
			<td align="center">������</td>
			<td align="center"><img src="/Images/Btn/icon_02.gif" style="cursor:hand;" align="absmiddle"></td>
			<td align="center"><img src="/Images/Btn/icon_02.gif" style="cursor:hand;" align="absmiddle"></td>
			<td align="center"><img src="/Images/Btn/icon_02.gif" style="cursor:hand;" align="absmiddle"></td>
			<td align="center">
				<img src="/Images/Btn/BtnIconModify.gif" title="����" style="cursor:hand;" align="absmiddle" onclick="javascript:fn_update('1111','UP');">
			</td>
	</tr>
	<tr id="cTR1" style="cursor:hand;" bgcolor="#ffffff" >

			<td align="center">2</td>
			<td align="center">0000000000</td>
			<td align="center">�������-����-������,���α�������</td>
			<td align="center"><img src="/Images/Btn/icon_02.gif" style="cursor:hand;" align="absmiddle"></td>
			<td align="center">2009-01-11</td>
			<td align="center">���������庴��</td>
			<td align="center">����</td>
			<td align="center">������</td>
			<td align="center"><img src="/Images/Btn/icon_02.gif" style="cursor:hand;" align="absmiddle"></td>
			<td align="center"><img src="/Images/Btn/icon_02.gif" style="cursor:hand;" align="absmiddle"></td>
			<td align="center"><img src="/Images/Btn/icon_03.gif" style="cursor:hand;" align="absmiddle"></td>
			<td align="center">
				<img src="/Images/Btn/BtnIconModify.gif" title="����" style="cursor:hand;" align="absmiddle" onclick="javascript:fn_update('1111','UP');">
			</td>
	</tr>
	<tr id="cTR1" style="cursor:hand;" bgcolor="#ffffff" >

			<td align="center">3</td>
			<td align="center">0000000000</td>
			<td align="center">�������-����-������,���α�������</td>
			<td align="center"><img src="/Images/Btn/icon_02.gif" style="cursor:hand;" align="absmiddle"></td>
			<td align="center">2009-02-20</td>
			<td align="center">���������庴��</td>
			<td align="center">����</td>
			<td align="center">������</td>
			<td align="center"><img src="/Images/Btn/icon_02.gif" style="cursor:hand;" align="absmiddle"></td>
			<td align="center"><img src="/Images/Btn/icon_02.gif" style="cursor:hand;" align="absmiddle"></td>
			<td align="center"><img src="/Images/Btn/icon_02.gif" style="cursor:hand;" align="absmiddle"></td>
			<td align="center">
				<img src="/Images/Btn/BtnIconModify.gif" title="����" style="cursor:hand;" align="absmiddle" onclick="javascript:fn_update('1111','UP');">
			</td>
	</tr>


</table>

<table border="0" cellpadding="0" cellspacing="0" width="1200" align="center">
	<tr><td height="5"></td></tr>
	<tr><td height="1" bgcolor="#D6D6D6"></td></tr>
	<tr height="22" bgcolor="#EEF6FF"><td align="center">1  2  3  4  5  6</td></tr>
	<tr><td height="1" bgcolor="#D6D6D6"></td></tr>
</table>
<table border="0" cellpadding="0" cellspacing="0" width="1200" align="center">
	<tr><td height="5"></td></tr>
</table>

<table border="0" cellpadding="0" cellspacing="0" width="1200" align="center">
	<tr><td><iframe frameborder=0 marginheight=0 marginwidth=0 width="100%" height="280" scrolling="no" name="AsInfo1fr"></iframe></td></tr>
</table>




<table border="0" width="1200" cellpadding="0" cellspacing="3" bgcolor="#EFECE5" align="center">
	<tr height="60">
		<td width=150 class="type-prov"><div align='left'>���ǳ�����ȣ:&nbsp;<input name="txtCid" type="text"class="type-prov" style="border-width:0px ; border-color:#cccccc ; border-style:solid" size="4" value ="1234" readonly></div></td>
		<td width=150 class="type-prov"><div align='left'>�߽Ź�ȣ:&nbsp;<input name="txtCid" type="text"class="type-prov" style="border-width:0px ; border-color:#cccccc ; border-style:solid" size="12" value ="01051850478" readonly></div></td>
		
		
		<td width=150 class="type-prov"><div align='right'>��ȭ�ɱ�:&nbsp;<input name="txtTelno" type="text"class="type-prov" style="border-width:1px ; border-color:#cccccc ; border-style:solid" size="12"></div></td>
		<td ><img align=ABSBOTTOM id="��ȭ�ɱ�" Style="cusor:hand;" src="/Images/Btn/BtnCallSend.gif" border="0" onclick=vfn_MakeCall() Style="cusor:hand;" >
		<img align=ABSBOTTOM id="��ȭ�ޱ�" src="/Images/Btn/BtnCallGet.gif" border="0" onclick="javascript:fn_PickUp();" Style="cusor:hand;">
		<img align=ABSBOTTOM id="��ȭ����" src="/Images/Btn/BtnCallOut.gif" border="0" onclick=vfn_Disconnect() Style="cusor:hand;">
		</td><td align="right" ><span class="blue_bold">��������</span>:<input name="txtStatus" type="text" class="type-prov" style="border-width:1px ; border-color:#cccccc ; border-style:none" size="9" ondblclick=vfn_Pickup() value ="�����" readonly><span class="blue_bold">�ذ���ð�</span>:<input name="txtStatus" type="text" class="type-prov" style="border-width:1px ; border-color:#cccccc ; border-style:none" size="9" ondblclick=vfn_Pickup() value ="00:05:15" readonly>
		<img id="BtnReady" src="/Images/Btn/BtnReady.gif" Style="cusor:hand;" align="absmiddle" onclick="vfn_SetAgentStatus();">
		<!--<select name="cboSetStatus" class="type-prov" onchange="vfn_SetAgentStatus()">
		  <option value=''>���º���</option>
		  <option value='Ready'>Ready</option>
		  <option value='01'>Ÿ����</option>
		  <option value='02'>�Ļ�</option>
		  <option value='03'>�̼�</option>
		  <option value='04'>�޽�</option>
					</select></div>--></td>  
		
	</tr>
</table>
<!-- #include virtual="/Include/Bottom.asp" -->