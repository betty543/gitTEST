<!-- #include virtual="/Include/Top.asp" -->

<%
'####### �Ķ���� ##################################################################################
SS_Login_Secgroup = SESSION("SS_Login_Secgroup")
SS_Login_Grade = SESSION("SS_Login_Grade")
SS_Login_CTIID = SESSION("SS_Login_CTIID")
SS_Login_EXTNO = SESSION("SS_Login_EXTNO")
SS_LoginID = SESSION("SS_LoginID")

QueryYN = request("QueryYN")
FromDate = request("FromDate")
ToDate = request("ToDate")
whereCD1 = Trim(request("whereCD1"))
whereCD2 = Trim(request("whereCD2"))
whereCD3 = Trim(request("whereCD3"))
whereCD7 = Trim(request("whereCD7"))
whereCD8 = Trim(request("whereCD8"))
whereCD9 = Trim(request("whereCD9"))

whereCD2 = Replace(whereCD2," ","")

CHANNELGB1 = request("CHANNELGB1")
CHANNELGB2 = request("CHANNELGB2")
CHANNELGB3 = request("CHANNELGB3")
CHANNELGB4 = request("CHANNELGB4")

MAN = request("MAN")
WOMAN = request("WOMAN")

If QueryYN = "" Then
	whereCD3 = "1"
End If



if FromDate = "" then FromDate = left(Date(),7) & "-01" end If
if ToDate = "" then ToDate = date() end If

pageWHERE = "QueryYN=" & QueryYN & "&FromDate=" & FromDate & "&ToDate=" & ToDate
pageWHERE = pageWHERE & "&whereCD1=" & whereCD1 & "&whereCD2=" & whereCD2 & "&whereCD3=" & whereCD3 & "&whereCD7=" & whereCD7 & "&whereCD8=" & whereCD8 & "&whereCD9=" & whereCD9
pageWHERE = pageWHERE & "&channelGb1=" & CHANNELGB1 & "&channelGb2=" & CHANNELGB2 & "&channelGb3=" & CHANNELGB3 & "&channelGb4=" & CHANNELGB4& "&MAN="&MAN& "&WOMAN="&WOMAN

'RESPONSE.WRITE pageWHERE

If CHANNELGB1 <> "" then
	CHANNELGB = "''" & CHANNELGB1 & "''"
End If
If CHANNELGB2 <> "" And CHANNELGB = "" then
	CHANNELGB = "''" & CHANNELGB2 & "''"
ElseIf CHANNELGB2 <> "" then
	CHANNELGB = CHANNELGB & ",''" & CHANNELGB2 & "''"
End If
If CHANNELGB3 <> "" And CHANNELGB = "" then
	CHANNELGB = "''" & CHANNELGB3 & "''"
ElseIf CHANNELGB3 <> "" then
	CHANNELGB = CHANNELGB & ",''" & CHANNELGB3 & "''"
End If
If CHANNELGB4 <> "" And CHANNELGB = "" then
	CHANNELGB = "''" & CHANNELGB4 & "''"
ElseIf CHANNELGB4 <> "" then
	CHANNELGB = CHANNELGB & ",''" & CHANNELGB4 & "''"
End If

JEONDOR = ""
If MAN = "" Then
	JEONDOR = "N"
Else
	JEONDOR = "Y"
End If
If WOMAN = "" Then
	JEONDOR = JEONDOR & "N"
Else
	JEONDOR = JEONDOR & "Y"
End if

%>

<div id='CalendarLayer' style='display:none;'><iframe name='CalendarFrame' src='/include/Calendar.asp' width='172' height='177' border='0' frameborder='0' scrolling='no'></iframe></div>

<script>
	function fn_Search(){
		if (document.all.whereCD8.value == document.all.whereCD9.value){
			alert('�����׸�� �����׸��� �ٸ��� �����ϼ���!')
			return false;
		}
		document.inUpFrm.QueryYN.value = "Y";
		document.inUpFrm.submit();
	}

	function fn_Xls() {
		location.href="list01_Xls.asp?<%=pageWHERE%>";
	}
</script>

<table border="0" width="1200" cellpadding="0" cellspacing="3" bgcolor="#EFECE5" align="center">
	<tr bgcolor="#FFFFFF">
		<td>

			<form name="inUpFrm" method="post" action="<%=Menu_2nd%>" onsubmit="return fn_Search(this);" style="margin:0">
				<input type="hidden" name="QueryYN" value="<%=QueryYN%>">

				<table width="100%" border="0" cellspacing="1" cellpadding="0" style="border:#E1DED6 solid 1px">
					<tr>
						<td class="TDCont" colspan="7">
							��ȭ���� :
							<input type="checkbox" name="CHANNELGB1" <% if CHANNELGB1 = "0179" then %>checked<% end if %> value="0179" class="none" />080����
							<input type="checkbox" name="CHANNELGB2" <% if CHANNELGB2 = "13031" then %>checked<% end if %> value="13031" class="none"  >1303����
							<input type="checkbox" name="CHANNELGB3" <% if CHANNELGB3 = "13032" then %>checked<% end if %> value="13032" class="none"  >1303������
							<input type="checkbox" name="CHANNELGB4" <% if CHANNELGB4 = "13033" then %>checked<% end if %> value="13033" class="none"  >1303������
							<!--<input type="checkbox" name="CHANNELGB4" <% if CHANNELGB4 = "CYBER" then %>checked<% end if %> value="CYBER" class="none"  >���̹�-->
						</td>

					</tr>
					<tr>
						<td class="TDCont" >��ȸ�Ⱓ :
							<input value="<%=FromDate%>" name="FromDate" readonly type="text" size="10" onfocus="setFocusColor(this); new CalendarFrame.Calendar(this);" onblur="setOutColor(this);">
							~
							<input value="<%=ToDate%>" name="ToDate" readonly type="text" size="10" onfocus="setFocusColor(this); new CalendarFrame.Calendar(this);" onblur="setOutColor(this);">
						</td>
						<td class="TDCont" COLSPAN = '6'>������� :

							<%
							'======= ó������ �ڵ� �������� ==================================================
							SqlCode = "SELECT BCLASS CODE, CLASSNAME CODENAME FROM TB_ARMYINFO"
							SqlCode = SqlCode& " WHERE ACLASS = 'Q' AND BCLASS IS NOT NULL AND CCLASS IS NULL"
							SqlCode = SqlCode& " ORDER BY ACLASS"
							set RsCode = db.execute(SqlCode)

							Do Until rsCode.eof
								sCode = RsCode("CODE")
								sCodeName = RsCode("CODENAME")
								If InStr(whereCD2,sCode) > 0 then
									sChecked = "checked"
								ElseIf whereCD2 = "" Then
									sChecked = ""
								else
									sChecked = ""
								End if
								%>
								<input type="checkbox" name="whereCD2" value="<%=sCode%>" class="none" <%=sChecked%>><%=sCodeName%>
								<%
								rsCode.movenext
							loop
							%>
							&nbsp;

						</td>

						<td class="TDR5px" rowspan='2'>
							<img src="/Images/Btn/BtnSearch.gif" align="absmiddle" style="cursor:hand;" onClick="fn_Search();">
							<img src="/Images/Btn/BtnExcel.gif" align="absmiddle" style="cursor:hand;" onClick="fn_Xls();">
						</td>
					</tr>
					<tr>
						<td class="TDCont" WIDTH="300">���α����׸� : &nbsp;
							<select name="whereCD8" size="1" class="ComboFFFCE7" >
								<%=printSelect("����","����","" &whereCD8& "")%>
								<%=printSelect("�ð�","�ð�","" &whereCD8& "")%>
								<%=printSelect("����","����","" &whereCD8& "")%>
								<%=printSelect("�������","�������","" &whereCD8& "")%>
								<%=printSelect("���","���","" &whereCD8& "")%>
								<%=printSelect("�δ�1��","�δ�1��","" &whereCD8& "")%>
								<%=printSelect("�δ�2��","�δ�2��","" &whereCD8& "")%>
								<%=printSelect("�δ�3��","�δ�3��","" &whereCD8& "")%>
								<%=printSelect("�δ�4��","�δ�4��","" &whereCD8& "")%>
								<%=printSelect("�δ�","�δ�","" &whereCD8& "")%>
								<%=printSelect("���о�","���о�","" &whereCD8& "")%>
								<%=printSelect("��ȭ�ð�","��ȭ�ð�","" &whereCD8& "")%>
								<%=printSelect("��ġ��","��ġ��","" &whereCD8& "")%>
								<%=printSelect("������","������","" &whereCD8& "")%>
								<%=printSelect("�������","�������","" &whereCD8& "")%>
								<%=printSelect("��������","��������","" &whereCD8& "")%>
								<%=printSelect("����������","����������","" &whereCD8& "")%>
								<%=printSelect("��","��","" &whereCD8& "")%>
							</select>
						</td>
						<td class="TDCont" WIDTH="200">�����׸� : &nbsp;
							<select name="whereCD9" size="1" class="ComboFFFCE7" >
								<%=printSelect("�������","�������","" &whereCD9& "")%>
								<%=printSelect("���","���","" &whereCD9& "")%>
								<%=printSelect("�δ�1��","�δ�1��","" &whereCD9& "")%>
								<%=printSelect("�δ�2��","�δ�2��","" &whereCD9& "")%>
								<!--<%=printSelect("�δ�3��","�δ�3��","" &whereCD9& "")%>	-->
								<!--<%=printSelect("�δ�4��","�δ�4��","" &whereCD9& "")%>
								<%=printSelect("�δ�","�δ�","" &whereCD9& "")%>	-->
								<%=printSelect("���о�","���о�","" &whereCD9& "")%>
								<%=printSelect("��ȭ�ð�","��ȭ�ð�","" &whereCD9& "")%>
								<%=printSelect("��ġ��","��ġ��","" &whereCD9& "")%>
								<%=printSelect("������","������","" &whereCD9& "")%>
								<%=printSelect("�������","�������","" &whereCD9& "")%>
								<%=printSelect("��������","��������","" &whereCD9& "")%>
								<%=printSelect("����������","����������","" &whereCD9& "")%>
								<%=printSelect("����","����","" &whereCD9& "")%>
								<%=printSelect("�ð�","�ð�","" &whereCD9& "")%>
								<%=printSelect("����","����","" &whereCD9& "")%>
							</select>
						</td>
						<td class="TDCont" WIDTH="200">���� :
							<%
							'======= ���� �������� ==================================================
							SqlCode = "SELECT USERID, USERNAME FROM TB_USERINFO"
							SqlCode = SqlCode& " WHERE USEYN='Y' "
							SqlCode = SqlCode& " AND SECGROUP = 'A'"
							if SS_Login_Grade <> "A" then
								SqlCode = SqlCode& "	AND GRADE = '"&SS_Login_Grade&"'"
							end if
							if SS_Login_Secgroup = "A" then	'�����϶��� ���͸�
								SqlCode = SqlCode& "	AND USERID = '" &SS_LoginID&"'"
							end if

							SqlCode = SqlCode& " ORDER BY USEYN DESC, GRADE ASC, USERNAME ASC"
							set RsCode = db.execute(SqlCode)
							%>
							<select name="whereCD1" size="1" class="ComboFFFCE7">
								<option value="">��ü</option>
								<%
								IF NOT(RsCode.Eof OR RsCode.bof) THEN
									DO until RsCode.EOF
										CODE = RsCode("USERID")
										CODENAME = RsCode("USERNAME")
										%>
										<%=printSelect("" &CODENAME& "","" &CODE& "","" &whereCD1& "")%>
										<%
										RsCode.MoveNext
									LOOP
								END IF
								RsCode.Close
								set RsCode = NOTHING

								'======= ���� �������� ==================================================
								SqlCode = "SELECT CTIID, USERNAME FROM TB_USERINFO"
								SqlCode = SqlCode& " WHERE USEYN='N'  and	outdate >= '"&DateAdd("d",1,DateAdd("m",-1,Date())) &"'"
								if SS_Login_Grade <> "A" then
									SqlCode = SqlCode& "	AND GRADE = '"&SS_Login_Grade&"'"
								end if
								if SS_Login_Secgroup = "A" then	'�����϶��� ���͸�
									SqlCode = SqlCode& "	AND USERID = '" &SS_LoginID&"'"
								end if

								SqlCode = SqlCode& " ORDER BY USEYN DESC, GRADE ASC, USERNAME ASC"
								set RsCode = db.execute(SqlCode)

								IF NOT(RsCode.Eof OR RsCode.bof) THEN
									DO until RsCode.EOF
										CODE = RsCode("CTIID")
										CODENAME = "[����]"&RsCode("USERNAME")
										%>
										<%=printSelect("" &CODENAME& "","" &CODE& "","" &whereCD1& "")%>
										<%
										RsCode.MoveNext
									LOOP
								END IF
								RsCode.Close
								set RsCode = NOTHING
								%>
							</select>
						</td>
						<td class="TDCont" colspan="1">
							��&nbsp;&nbsp;&nbsp;&nbsp;�� : 
							<input type="checkbox" name="MAN" <% if MAN = "Y" then %>checked<% end if %> value="Y" class="none"  >����
							<input type="checkbox" name="WOMAN" <% if WOMAN = "Y" then %>checked<% end if %> value="Y" class="none"  >����
							<!--<input type="checkbox" name="CHANNELGB4" <% if CHANNELGB4 = "CYBER" then %>checked<% end if %> value="CYBER" class="none"  >���̹�-->
						</td>
					</tr>
				</table>
			</form>
		</td>
	</tr>
</table>

<table border="0" width="1000" cellpadding="0" cellspacing="0" align="center"><tr height="5"><td></td></tr></table>

<% If QueryYN = "Y" Then %>

	<table border="0" cellpadding="0" cellspacing="0" align="center">
		<tr>
			<td>
				<%
				'�����׸�
				cols = 1
				sDepth = 1
				'response.write whereCD8
				', ����,�������, ��������, ����������, ���� , �ð�, ����,��ȭ�ð�, - 1depth
				if whereCD8 = "�������" then'������� - 2depth
					sSero = "<td align='center' class='TDCont'  width='300' colspan='2' "
					sSero1 = "�������</td>"
					cols = 2
					sDepth = 2
					sSelectCol = "CHANNELGB_B col_1, CHANNELGB_C col_2"
					sGroupBy = "CHANNELGB_B, CHANNELGB_C"
					sSelectCol1 = "CHANNELGB_B, CHANNELGB_C"
					sNullCol = " AND ( isnull(CHANNELGB_B,'''') <> '''' or isnull(CHANNELGB_C,'''') <> '''') "
				elseif whereCD8 = "���" then'���     - 3depth
					sSero = "<td align='center' class='TDCont'  width='450' colspan='3' "
					sSero1 = "���</td>"
					cols = 3
					sDepth = 3
					sSelectCol = "LEVEL_B col_1, LEVEL_C col_2, LEVEL_D col_3"
					sGroupBy = "LEVEL_B, LEVEL_C, LEVEL_D"
					sSelectCol1 = "LEVEL_B, LEVEL_C, LEVEL_D"
'					sSelectCol1 = "ISNULL(LEVEL_B,'''') LEVEL_B, ISNULL(LEVEL_C,'''') LEVEL_C, ISNULL(LEVEL_D,'''') LEVEL_D"
					sNullCol = " AND ( RTRIM(isnull(LEVEL_B,'''')) <> '''' or RTRIM(isnull(LEVEL_C,'''')) <> '''' or RTRIM(isnull(LEVEL_D,'''')) <> '''') "
					sNullCol = " AND RTRIM(isnull(LEVEL_B,'''')) <> ''''"
				elseif whereCD8 = "�δ�" then'�δ�	  - 5depth
					sSero = "<td align='center' class='TDCont'  width='750' colspan='5' "
					sSero1 = "�δ�</td>"
					cols = 5
					sDepth = 5
					sSelectCol = "SOSOKGB_A col_1, SOSOKGB_B col_2, SOSOKGB_C col_3, SOSOKGB_D col_4, SOSOKGB_E col_5"
					sGroupBy = "SOSOKGB_A, SOSOKGB_B, SOSOKGB_C, SOSOKGB_D, SOSOKGB_E"
					sSelectCol1 = "SOSOKGB_A, SOSOKGB_B, SOSOKGB_C, SOSOKGB_D, SOSOKGB_E"
					sNullCol = " AND ( isnull(SOSOKGB_A,'''') <> '''' or isnull(SOSOKGB_B,'''') <> '''' or isnull(SOSOKGB_C,'''') <> '''' or isnull(SOSOKGB_D,'''') <> '''' or isnull(SOSOKGB_E,'''') <> '''') "
				elseif whereCD8 = "�δ�1��" then'�δ�	  - 5depth
					sSero = "<td align='center' class='TDCont'  width='750' colspan='1' "
					sSero1 = "�δ�</td>"
					cols = 1
					sDepth = 1
					sSelectCol = "SOSOKGB_A col_1"
					sGroupBy = "SOSOKGB_A"
					sSelectCol1 = "SOSOKGB_A"
					sNullCol = " AND isnull(SOSOKGB_A,'''') <> '''' "
				elseif whereCD8 = "�δ�2��" then'�δ�	  - 5depth
					sSero = "<td align='center' class='TDCont'  width='750' colspan='2' "
					sSero1 = "�δ�</td>"
					cols = 2
					sDepth = 2
					sSelectCol = "SOSOKGB_A col_1, SOSOKGB_B col_2"
					sGroupBy = "SOSOKGB_A, SOSOKGB_B"
					sSelectCol1 = "SOSOKGB_A, SOSOKGB_B"
					sNullCol = " AND ( isnull(SOSOKGB_A,'''') <> '''' or isnull(SOSOKGB_B,'''') <> '''') "
				elseif whereCD8 = "�δ�3��" then'�δ�	  - 5depth
					sSero = "<td align='center' class='TDCont'  width='750' colspan='3' "
					sSero1 = "�δ�</td>"
					cols = 3
					sDepth = 3
					sSelectCol = "SOSOKGB_A col_1, SOSOKGB_B col_2, SOSOKGB_C col_3"
					sGroupBy = "SOSOKGB_A, SOSOKGB_B, SOSOKGB_C"
					sSelectCol1 = "SOSOKGB_A, SOSOKGB_B, SOSOKGB_C"
					sNullCol = " AND ( isnull(SOSOKGB_A,'''') <> '''' or isnull(SOSOKGB_B,'''') <> '''' or isnull(SOSOKGB_C,'''') <> '''') "
				elseif whereCD8 = "�δ�4��" then'�δ�	  - 5depth
					sSero = "<td align='center' class='TDCont'  width='750' colspan='4' "
					sSero1 = "�δ�</td>"
					cols = 4
					sDepth = 4
					sSelectCol = "SOSOKGB_A col_1, SOSOKGB_B col_2, SOSOKGB_C col_3, SOSOKGB_D col_4"
					sGroupBy = "SOSOKGB_A, SOSOKGB_B, SOSOKGB_C, SOSOKGB_D"
					sSelectCol1 = "SOSOKGB_A, SOSOKGB_B, SOSOKGB_C, SOSOKGB_D"
					sNullCol = " AND ( isnull(SOSOKGB_A,'''') <> '''' or isnull(SOSOKGB_B,'''') <> '''' or isnull(SOSOKGB_C,'''') <> '''' or isnull(SOSOKGB_D,'''') <> '''') "
				elseif whereCD8 = "���о�" then'���о� - 2depth
					sSero = "<td align='center' class='TDCont'  width='300' colspan='2' "
					sSero1 = "���о�</td>"
					cols = 2
					sDepth = 2
					sSelectCol = "CALLCLASS_B col_1, CALLCLASS_C col_2"
					sGroupBy = "CALLCLASS_B, CALLCLASS_C"
					sSelectCol1 = "CALLCLASS_B, CALLCLASS_C"
					sNullCol = " AND ( isnull(CALLCLASS_B,'''') <> '''' or isnull(CALLCLASS_C,'''') <> '''') "
				elseif whereCD8 = "��ġ��" then'��ġ�� - 1depth
					sSero = "<td align='center' class='TDCont'  width='300' colspan='2' "
					sSero1 = "��ġ��</td>"
					cols = 2
					sDepth = 2
					sSelectCol = "PROCESSGB_B col_1, PROCESSGB_C col_2"
					sGroupBy = "PROCESSGB_B, PROCESSGB_C"
					sSelectCol1 = "PROCESSGB_B, PROCESSGB_C"
					sNullCol = " AND ( isnull(PROCESSGB_B,'''') <> '''' or isnull(PROCESSGB_C,'''') <> '''') "
				elseif whereCD8 = "������" then'���� - 1depth
					sSero = "<td align='center' class='TDCont'  width='150' colspan='1' "
					sSero1 = "������</td>"
					sSelectCol = "WEATHER col_1"
					sGroupBy = "WEATHER"
					sSelectCol1 = "WEATHER"
					sNullCol = " AND  isnull(WEATHER,'''') <> '''' "
				elseif whereCD8 = "�������" then'������� - 1depth
					sSero = "<td align='center' class='TDCont'  width='150' colsFpan='1' "
					sSero1 = "�������</td>"
					sSelectCol = "CALLFLAG col_1"
					sGroupBy = "CALLFLAG"
					sSelectCol1 = "CALLFLAG"
					sNullCol = " AND  isnull(CALLFLAG,'''') <> '''' "
				elseif whereCD8 = "��������" then'�������� - 1depth
					sSero = "<td align='center' class='TDCont'  width='150' colspan='1' "
					sSero1 = "��������</td>"
					sSelectCol = "FAMILYGB col_1"
					sGroupBy = "FAMILYGB"
					sSelectCol1 = "FAMILYGB"
					sNullCol = " AND  isnull(FAMILYGB,'''') <> '''' "
				elseif whereCD8 = "����������" then'���������� - 1depth
					sSero = "<td align='center' class='TDCont'  width='150' colspan='1' "
					sSero1 = "����������</td>"
					sSelectCol = "CALLKIND_B col_1"
					sGroupBy = "CALLKIND_B"
					sSelectCol1 = "CALLKIND_B"
					sNullCol = " AND  isnull(CALLKIND_B,'''') <> '''' "
				elseif whereCD8 = "����" then'���� - 1depth
					sSero = "<td align='center' class='TDCont'  width='150' colspan='1' "
					sSero1 = "����</td>"
					sSelectCol = "INCODE col_1"
					sGroupBy = "INCODE"
					sSelectCol1 = "INCODE"
					sNullCol = " AND  isnull(INCODE,'''') <> '''' "
				elseif whereCD8 = "�ð�" then'�ð� - 1depth
					sSero = "<td align='center' class='TDCont'  width='150' colspan='1' "
					sSero1 = "�ð�</td>"
					sSelectCol = "datepart(hour,JUBTIME) COL_1"
					sSelectCol1 = "JUBTIME"
					sGroupBy = "datepart(hour,JUBTIME)"
					sNullCol = " "
				elseif whereCD8 = "����" then'���� - 1depth
					sSero = "<td align='center' class='TDCont'  width='150' colspan='1' "
					sSero1 = "����</td>"
					sSelectCol = "datepart(WEEKDAY,JUBTIME) COL_1"
					sGroupBy = "datepart(WEEKDAY,JUBTIME)"
					sSelectCol1 = "JUBTIME"
					sNullCol = " "
				elseif whereCD8 = "��ȭ�ð�" then'��ȭ�ð� - 1depth
					sSero = "<td align='center' class='TDCont'  width='150' colspan='1' "
					sSero1 = "��ȭ�ð�</td>"
					sSelectCol = "CALLTIME COL_1"
					sSelectCol1 = "CALLTIME"
					sGroupBy = "CALLTIME"
					sNullCol = " "
				elseif whereCD8 = "��" then'�� - 1depth
					sSero = "<td align='center' class='TDCont'  width='150' colspan='1' "
					sSero1 = "��</td>"
					sSelectCol = "convert(varchar(7),JUBTIME,121) COL_1"
					sGroupBy = "convert(varchar(7),JUBTIME,121)"
					sSelectCol1 = "JUBTIME"
					sNullCol = " "
				end if

				'---- ����������
				'sCOLNM = "CALLKIND_B"

				sSQL = "DELETE FROM TMP_CODE_VALUE"
				db.execute(sSQL)
				', ����,�������, ��������, ����������, ���� , �ð�, ����,��ȭ�ð�, - 1depth
				if whereCD9 = "�������" then'������� - 2depth

					'-----�����׸� �Ѹ���
					rowspan = 2
					sSQL = "	select ACLASS, BCLASS, CCLASS, CLASSNAME from TB_ARMYINFO "
					sSQL = sSQL & "	where ACLASS = 'Q' AND BCLASS IS NOT NULL AND CCLASS IS NULL"
					sSQL = sSQL & "	ORDER BY ACLASS, BCLASS, CCLASS "
					set RsCode = db.execute(sSQL)

					Do Until rsCode.eof

						sCode = RsCode("BCLASS")
						sCodeName = RsCode("CLASSNAME")

						'2DEPTH �� ã��
						iCol = 0

						'secondLine = ""
						sCodeList = ""
						sSQL = "	select ACLASS, BCLASS, CCLASS, CLASSNAME from TB_ARMYINFO "
						sSQL = sSQL & "	where ACLASS = 'Q' AND BCLASS = '" &sCode&"'  AND CCLASS IS NOT NULL"
						sSQL = sSQL & "	ORDER BY ACLASS, BCLASS, CCLASS "
						set RsCode1 = db.execute(sSQL)

						Do Until rsCode1.eof

							sCode = RsCode1("CCLASS")
							sCodeName = RsCode1("CLASSNAME")

							iCol = iCol + 1
							cols = cols + 1

							sSQL = "INSERT INTO TMP_CODE_VALUE VALUES ("&cols& ",'CHANNELGB_C','" & sCode & "')"
							db.execute(sSQL)

							If sCodeList = "" then
								sCodeList = sCode
							Else
								sCodeList = sCodeList & "|" & sCode
							End if

							secondLine = secondLine & "<td align='center' class='TDCont'  width='150'>" & sCodeName & "</td>"
							rsCode1.movenext
							'�Ұ�
							
						Loop

						If iCol = 0 Then
							
							sSQL = "INSERT INTO TMP_CODE_VALUE VALUES ("&cols& ",'CHANNELGB_B','" & sCode & "')"
							db.execute(sSQL)
							cols = cols + 1

							firstLine = firstLine & "<td align='center' class='TDCont' colspan="&iCol+1&" rowspan='2' width='150'>"&RsCode("CLASSNAME")&"</td>"

						Else
							
							cols = cols + 1
							secondLine = secondLine & "<td align='center' class='TDCont'  width='150'>��</td>"
							sSQL = "INSERT INTO TMP_CODE_VALUE VALUES ("&cols& ",'CHANNELGB_C','" & sCodeList & "')"
							db.execute(sSQL)

							firstLine = firstLine & "<td align='center' class='TDCont' colspan="&iCol+1&">"&RsCode("CLASSNAME")&"</td>"

						End if

						rsCode.movenext
						'�Ұ�
					Loop
					'�Ѱ�
					%>
					
					<DIV style="OVERFLOW-X: auto;OVERFLOW-Y: auto; MARGIN: 0px 0px 0px 0px; WIDTH: 1200; HEIGHT=400;">
						<table width="<%=sWidth%>"  border="1" cellpadding="0" cellspacing="0" bgcolor="#EFECE5" align="center" bordercolor="black" bordercolordark="white" bordercolorlight="black">
							
							<%
							firstLine = firstLine & "<td align='center' class='TDCont' rowspan="& rowspan &" width='150'>��</td>"
							firstLine = "<tr bgcolor='#EEF6FF'>" & sSero & " rowspan="& rowspan &">"& sSero1 & firstLine &"</tr>"
							secondLine = "<tr bgcolor='#EEF6FF'>" & secondLine &"</tr>"
							response.write firstLine
							response.write secondLine
	
							''-----�����׸� �Ѹ���
							sCOLNM = "CHANNELGB_B"
							sCOLCD = ""
	
							sSQL = " EXEC SP_SUM_BY_HISTORY_BCLASS " & sDepth & ",'" & FromDate & "','" & ToDate & "','TB_LIFECALLHISTORY','"&whereCD2 & "','"& whereCD1 & "','"&sCOLNM&"','Q','','"&sGroupBy&"','"&sSelectCol&"','"&sSelectCol1&"','"&sNullCol&"','"&CHANNELGB&"','"&JEONDOR&"','',''"
							'response.write sSQL
							set rsSUM = db.execute(sSQL)
	
							firstLine = ""
	
							Do Until rsSUM.eof
								sBG = "#ffffff"
	
								firstLine = ""
	
								'--------------Ű�� �ش��ϴ� ��
								For i = 1 To sDepth
									sUser = rsSUM("col_"&i)
									'sCodeName = db_GetUserName(sUser)
	
									If IsNull(rsSUM("col_"&i)) Then
										
										sBG = "#EEF6FF"
	
										firstLine = firstLine & "<td align='center' class='TDCont' colspan='"&sDepth-i+1&"' WIDTH='200'>��</td>"
	
										Exit for
										
									Else
										
										If i = 1 Then
											sCodeName = db_GetSumColNm(whereCD8,i,rsSUM("col_"&i),"","","","")
										ElseIf i = 2 Then
											sCodeName = db_GetSumColNm(whereCD8,i,rsSUM("col_"&i-1),rsSUM("col_"&i),"","","")
										ElseIf i = 3 Then
											sCodeName = db_GetSumColNm(whereCD8,i,rsSUM("col_"&i-2),rsSUM("col_"&i-1),rsSUM("col_"&i),"","")
										ElseIf i = 4 Then
											sCodeName = db_GetSumColNm(whereCD8,i,rsSUM("col_"&i-3),rsSUM("col_"&i-2),rsSUM("col_"&i-1),rsSUM("col_"&i),"")
										Else
											sCodeName = db_GetSumColNm(whereCD8,i,rsSUM("col_"&i-4),rsSUM("col_"&i-3),rsSUM("col_"&i-2),rsSUM("col_"&i-1),rsSUM("col_"&i))
										End if
	
										firstLine = firstLine & "<td align='center' class='TDCont' colspan='1' WIDTH='200'>"&sCodeName&"</td>"
	
									End If
								Next
	
								'--------------�����׸��� summary
								For i = sDepth + 1 To rsSUM.Fields.count
	
									firstLine = firstLine & "<td align='center' class='TDCont' colspan='1' WIDTH='200'>"&rsSUM("col_"&i)&"</td>"
	
								Next
	
								firstLine = "<tr bgcolor='"&sBG&"'>" & firstLine
	
								response.write firstLine & "</tr>"
	
								rsSUM.movenext
								'�Ұ�
							Loop
							%>
							
						</table>
					</div>
	
					<%
	
				elseif whereCD9 = "���" then'���     - 3depth
	
					rowspan = 3
					sSQL = "	select  ACLASS, BCLASS, CCLASS, CLASSNAME from TB_ARMYINFO "
					sSQL = sSQL & "	where ACLASS = 'P' AND BCLASS IS NOT NULL AND CCLASS IS NULL AND DCLASS IS NULL"
					sSQL = sSQL & "	ORDER BY ACLASS, BCLASS, CCLASS "
					set RsCode = db.execute(sSQL)
	
					Do Until rsCode.eof
	
						sCode = RsCode("BCLASS")
						sCodeName = RsCode("CLASSNAME")
	
						'2DEPTH �� ã��
						iCol = 0
						sCodeList = ""
						sCodeList_C = ""
						sSQL = "	select ACLASS, BCLASS, CCLASS, CLASSNAME from TB_ARMYINFO "
						sSQL = sSQL & "	where ACLASS = 'P' AND BCLASS = '" &sCode&"'  AND CCLASS IS NOT NULL AND DCLASS IS NULL"
						sSQL = sSQL & "	ORDER BY ACLASS, BCLASS, CCLASS "
						set RsCode1 = db.execute(sSQL)
	
						'response.write sSQL & vbcrlf
						Do Until rsCode1.eof
	
							sCode = RsCode1("CCLASS")
							sCodeName = RsCode1("CLASSNAME")
							iCol = iCol + 1
							iCol1 = 0
	
							'3DEPTH �� ã��
							If sCodeList_C = "" then
								sCodeList_C = RsCode1("CCLASS")
							Else
								sCodeList_C = sCodeList_C & "|" & RsCode1("CCLASS")
							End if
	
							'secondLine = ""
							sCodeList = ""
							sSQL = "	select ACLASS, BCLASS, CCLASS, DCLASS, CLASSNAME from TB_ARMYINFO "
							sSQL = sSQL & "	where ACLASS = 'P' AND BCLASS = '" &RsCode("BCLASS")&"'  AND CCLASS = '" & RsCode1("CCLASS")&"' AND DCLASS IS NOT NULL"
							sSQL = sSQL & "	ORDER BY ACLASS, BCLASS, CCLASS, DCLASS"
							set RsCode2 = db.execute(sSQL)
							'-------------------------------------------------------------------------------------------------------------------------------------
	
							'response.write sSQL & vbcrlf
							Do Until rsCode2.eof
	
								iCol1 = iCol1 + 1
								cols = cols + 1
								iCol = iCol + 1
								sSQL = "INSERT INTO TMP_CODE_VALUE VALUES ("&cols& ",'LEVEL_D','" & RsCode2("DCLASS") & "')"
								db.execute(sSQL)
	
								If sCodeList = "" then
									sCodeList = RsCode2("DCLASS")
								Else
									sCodeList = sCodeList & "|" & RsCode2("DCLASS")
								End if
	
								threeLine = threeLine & "<td align='center' class='TDCont'  width='150'>" & RsCode2("CLASSNAME") & "</td>"
								rsCode2.movenext
								'�Ұ�
	
							Loop
	
							If iCol1 <= 0 Then
								
								cols = cols + 1
								sSQL = "INSERT INTO TMP_CODE_VALUE VALUES ("&cols& ",'LEVEL_C','" & RsCode1("CCLASS") & "')"
								db.execute(sSQL)
								secondLine = secondLine & "<td align='center' class='TDCont' colspan=1 rowspan='2' width='150'>"&RsCode1("CLASSNAME")&"</td>"
								'sCodeList_C = ""
								
							Else
								
								iCol1 = iCol1 + 1
								cols = cols + 1
								sSQL = "INSERT INTO TMP_CODE_VALUE VALUES ("&cols& ",'LEVEL_D','" & sCodeList & "')"
								db.execute(sSQL)
								'iCol = iCol + 1
								threeLine = threeLine & "<td align='center' class='TDCont'  width='150' rowspan='1'>��</td>"
								secondLine = secondLine & "<td align='center' class='TDCont' colspan="&iCol1&">"&RsCode1("CLASSNAME")&"</td>"
								
							End If
							
							iCol1 = 0
							'response.write iCol & "," & vbcrlf
							rsCode1.movenext
							'�Ұ�
	
						Loop
	
						'response.write iCol & "," & vbcrlf
					If iCol = 0 Then
						
						cols = cols + 1
						sSQL = "INSERT INTO TMP_CODE_VALUE VALUES ("&cols& ",'LEVEL_B','" & RsCode("BCLASS") & "')"
						db.execute(sSQL)
	
						sCodeList_C = ""
						firstLine = firstLine & "<td align='center' class='TDCont' colspan="&iCol&" rowspan='3' width='150'>"&RsCode("CLASSNAME")&"</td>"
						
					Else
						
						cols = cols + 1
						iCol = iCol + 1
						sSQL = "INSERT INTO TMP_CODE_VALUE VALUES ("&cols& ",'LEVEL_C','" & sCodeList_C & "')"
						db.execute(sSQL)
						secondLine = secondLine & "<td align='center' class='TDCont'  width='150' rowspan='2'>��</td>"
						firstLine = firstLine & "<td align='center' class='TDCont' colspan="&iCol&" >"&RsCode("CLASSNAME")&"</td>"
						iCol = 0
						
					End if
					
					rsCode.movenext
	
					'�Ұ�
				Loop
				'�Ѱ�
				sWidth = cols * 100
				%>
				
				<DIV style="OVERFLOW-X: auto;OVERFLOW-Y: auto; MARGIN: 0px 0px 0px 0px; WIDTH: 1200; HEIGHT=400;">
				<table width="<%=sWidth%>"  border="1" cellpadding="0" cellspacing="0" bgcolor="#EFECE5" align="center" bordercolor="black" bordercolordark="white" bordercolorlight="black">
					
				<%
				firstLine = firstLine & "<td align='center' class='TDCont' rowspan="& rowspan &" width='150'>��</td>"
				firstLine = "<tr bgcolor='#EEF6FF'>" & sSero & " rowspan="& rowspan &">"& sSero1 & firstLine &"</tr>"
				secondLine = "<tr bgcolor='#EEF6FF'>" & secondLine &"</tr>"
				threeLine = "<tr bgcolor='#EEF6FF'>" & threeLine &"</tr>"
				response.write firstLine
				response.write secondLine
				response.write threeLine
	
				''-----�����׸� �Ѹ���
				sCOLNM = "LEVEL_B"
				sCOLCD = ""
	
				sSQL = " EXEC SP_SUM_BY_HISTORY_BCLASS " & sDepth & ",'" & FromDate & "','" & ToDate & "','TB_LIFECALLHISTORY','"&whereCD2 & "','"& whereCD1 & "','"&sCOLNM&"','Q','','"&sGroupBy&"','"&sSelectCol&"','"&sSelectCol1&"','"&sNullCol&"','"&CHANNELGB&"','"&JEONDOR&"','',''"
	
				'response.write sSQL

				'response.end
	
				set rsSUM = db.execute(sSQL)
				firstLine = ""
	
				Do Until rsSUM.eof
					sBG = "#ffffff"
					firstLine = ""
					'--------------Ű�� �ش��ϴ� ��
					For i = 1 To sDepth
						sUser = rsSUM("col_"&i)
						'sCodeName = db_GetUserName(sUser)
	
						If IsNull(rsSUM("col_"&i)) Then
							sBG = "#EEF6FF"
							firstLine = firstLine & "<td align='center' class='TDCont' colspan='"&sDepth-i+1&"' WIDTH='200'>��</td>"
							Exit for
						Else
							If i = 1 Then
								sCodeName = db_GetSumColNm(whereCD8,i,rsSUM("col_"&i),"","","","")
							ElseIf i = 2 Then
								sCodeName = db_GetSumColNm(whereCD8,i,rsSUM("col_"&i-1),rsSUM("col_"&i),"","","")
							ElseIf i = 3 Then
								sCodeName = db_GetSumColNm(whereCD8,i,rsSUM("col_"&i-2),rsSUM("col_"&i-1),rsSUM("col_"&i),"","")
							ElseIf i = 4 Then
								sCodeName = db_GetSumColNm(whereCD8,i,rsSUM("col_"&i-3),rsSUM("col_"&i-2),rsSUM("col_"&i-1),rsSUM("col_"&i),"")
							Else
								sCodeName = db_GetSumColNm(whereCD8,i,rsSUM("col_"&i-4),rsSUM("col_"&i-3),rsSUM("col_"&i-2),rsSUM("col_"&i-1),rsSUM("col_"&i))
							End if
							firstLine = firstLine & "<td align='center' class='TDCont' colspan='1' WIDTH='200'>"&sCodeName&"</td>"
						End If
					Next
					'--------------�����׸��� summary
					For i = sDepth + 1 To rsSUM.Fields.count
	
						firstLine = firstLine & "<td align='center' class='TDCont' colspan='1' WIDTH='200'>"&rsSUM("col_"&i)&"</td>"
	
					Next
					firstLine = "<tr bgcolor='"&sBG&"'>" & firstLine
					response.write firstLine & "</tr>"
					rsSUM.movenext
	
					'�Ұ�
				Loop
	
				%>
	
				</table>
				</div>
				
				<%
	
			elseif whereCD9 = "�δ�1��" then'�δ�	  - 5depth
	
				rowspan = 1
				sSQL = "	select  ACLASS, BCLASS, CCLASS, CLASSNAME from TB_ARMYINFO "
				sSQL = sSQL & "	where ACLASS < 'O' AND BCLASS IS NULL"
				sSQL = sSQL & "	ORDER BY ACLASS, BCLASS, CCLASS "
				set RsCode = db.execute(sSQL)
	
				Do Until rsCode.eof
	
					sCode = RsCode("ACLASS")
					sCodeName = RsCode("CLASSNAME")
	
					'2DEPTH �� ã��
					iCol = 0
					'secondLine = ""
					cols = cols + 1
					iCol = iCol + 1
					firstLine = firstLine & "<td align='center' class='TDCont' colspan=1 width='150'>"&RsCode("CLASSNAME")&"</td>"
					
					rsCode.movenext
					'�Ұ�
				Loop
				'�Ѱ�
				sWidth = cols * 150
				%>
				
				<DIV style="OVERFLOW-X: auto;OVERFLOW-Y: auto; MARGIN: 0px 0px 0px 0px; WIDTH: 1200; HEIGHT=400;">
				<table width="<%=sWidth%>"  border="1" cellpadding="0" cellspacing="0" bgcolor="#EFECE5" align="center" bordercolor="black" bordercolordark="white" bordercolorlight="black">
					
				<%
				firstLine = firstLine & "<td align='center' class='TDCont' rowspan="& rowspan &" width='150'>��</td>"
				firstLine = "<tr bgcolor='#EEF6FF'>" & sSero & " rowspan="& rowspan &">"& sSero1 & firstLine &"</tr>"
				response.write firstLine
	
				'---- �Ҽ�1��
				sCOLNM = "SOSOKGB_A"
				sCOLCD = ""
	
				sSQL = " EXEC SP_SUM_BY_HISTORY_ACLASS " & sDepth & ",'" & FromDate & "','" & ToDate & "','TB_LIFECALLHISTORY','"&whereCD2 & "','"& whereCD1 & "','"&sCOLNM&"','"&sCOLCD&"','"&sGroupBy&"','"&sSelectCol&"','"&sSelectCol1&"','"&sNullCol&"','"&CHANNELGB&"','"&JEONDOR&"','',''"
	
				'response.write sSQL
	
				set rsSUM = db.execute(sSQL)
				firstLine = ""
	
				Do Until rsSUM.eof
					sBG = "#ffffff"
					firstLine = ""
					For i = 1 To sDepth
						sUser = rsSUM("col_"&i)
						'sCodeName = db_GetUserName(sUser)
	
						If IsNull(rsSUM("col_"&i)) Then
							sBG = "#EEF6FF"
							firstLine = firstLine & "<td align='center' class='TDCont' colspan='"&sDepth-i+1&"' WIDTH='200'>��</td>"
							Exit for
						Else
							If i = 1 Then
								sCodeName = db_GetSumColNm(whereCD8,i,rsSUM("col_"&i),"","","","")
							ElseIf i = 2 Then
								sCodeName = db_GetSumColNm(whereCD8,i,rsSUM("col_"&i-1),rsSUM("col_"&i),"","","")
							ElseIf i = 3 Then
								sCodeName = db_GetSumColNm(whereCD8,i,rsSUM("col_"&i-2),rsSUM("col_"&i-1),rsSUM("col_"&i),"","")
							ElseIf i = 4 Then
								sCodeName = db_GetSumColNm(whereCD8,i,rsSUM("col_"&i-3),rsSUM("col_"&i-2),rsSUM("col_"&i-1),rsSUM("col_"&i),"")
							Else
								sCodeName = db_GetSumColNm(whereCD8,i,rsSUM("col_"&i-4),rsSUM("col_"&i-3),rsSUM("col_"&i-2),rsSUM("col_"&i-1),rsSUM("col_"&i))
							End if
							firstLine = firstLine & "<td align='center' class='TDCont' colspan='1' WIDTH='200'>"&sCodeName&"</td>"
						End If
					Next
	
					For i = sDepth + 1 To rsSUM.Fields.count
	
						firstLine = firstLine & "<td align='center' class='TDCont' colspan='1' WIDTH='200'>"&rsSUM("col_"&i)&"</td>"
	
					Next
					firstLine = "<tr bgcolor='"&sBG&"'>" & firstLine
					response.write firstLine & "</tr>"
					rsSUM.movenext
	
					'�Ұ�
				Loop
				%>
	
				</table>
				</div>
				
				<%
	
			elseif whereCD9 = "�δ�2��" then'�δ�	  - 5depth
	
				rowspan = 2
				sSQL = "	select  ACLASS, BCLASS, CCLASS, CLASSNAME from TB_ARMYINFO "
				sSQL = sSQL & "	where ACLASS < 'O' AND BCLASS IS NULL AND CCLASS IS NULL AND DCLASS IS NULL"
				sSQL = sSQL & "	ORDER BY ACLASS, BCLASS, CCLASS "
				set RsCode = db.execute(sSQL)
	
				Do Until rsCode.eof
	
					sCode = RsCode("ACLASS")
					sCodeName = RsCode("CLASSNAME")
	
					'2DEPTH �� ã��
					iCol = 0
					'secondLine = ""
					sCodeList = ""
					sSQL = "	select ACLASS, BCLASS, CCLASS, CLASSNAME from TB_ARMYINFO "
					sSQL = sSQL & "	where ACLASS = '" &sCode&"'  AND BCLASS IS NOT NULL AND CCLASS IS NULL"
					sSQL = sSQL & "	ORDER BY ACLASS, BCLASS, CCLASS "
					set RsCode1 = db.execute(sSQL)
	
					Do Until rsCode1.eof
	
						sCode = RsCode1("BCLASS")
						sCodeName = RsCode1("CLASSNAME")
						iCol = iCol + 1
	
						'3DEPTH �� ã��
	
						iCol1 = iCol1 + 1
						cols = cols + 1
						sSQL = "INSERT INTO TMP_CODE_VALUE VALUES ("&cols& ",'SOSOKGB_B','" & sCode & "')"
						db.execute(sSQL)
	
						If sCodeList = "" then
							sCodeList = sCode
						Else
							sCodeList = sCodeList & "|" & sCode
						End if
						secondLine = secondLine & "<td align='center' class='TDCont' colspan=1 width='150'>"&RsCode1("CLASSNAME")&"</td>"
						'response.write iCol & "," & vbcrlf
					
						rsCode1.movenext
						'�Ұ�
					Loop
	
				'response.write iCol & "," & vbcrlf
					If iCol = 0 Then
						sSQL = "INSERT INTO TMP_CODE_VALUE VALUES ("&cols& ",'SOSOKGB_A','" & sCode & "')"
						db.execute(sSQL)
						cols = cols + 1
						firstLine = firstLine & "<td align='center' class='TDCont' colspan="&iCol&" rowspan='2' width='150'>"&RsCode("CLASSNAME")&"</td>"
					Else
						sSQL = "INSERT INTO TMP_CODE_VALUE VALUES ("&cols& ",'SOSOKGB_B','" & sCodeList & "')"
						db.execute(sSQL)
						cols = cols + 1
						iCol = iCol + 1
						secondLine = secondLine & "<td align='center' class='TDCont'  width='150' rowspan='1'>��</td>"
						firstLine = firstLine & "<td align='center' class='TDCont' colspan="&iCol&" >"&RsCode("CLASSNAME")&"</td>"
						iCol = 0
					End if
					
					rsCode.movenext
	
					'�Ұ�
				Loop
				'�Ѱ�
				sWidth = cols * 150
				%>
				
				<DIV style="OVERFLOW-X: auto;OVERFLOW-Y: auto; MARGIN: 0px 0px 0px 0px; WIDTH: 1200; HEIGHT=400;">
				<table width="<%=sWidth%>"  border="1" cellpadding="0" cellspacing="0" bgcolor="#EFECE5" align="center" bordercolor="black" bordercolordark="white" bordercolorlight="black">
					
				<%
				firstLine = firstLine & "<td align='center' class='TDCont' rowspan="& rowspan &" width='150'>��</td>"
				firstLine = "<tr bgcolor='#EEF6FF'>" & sSero & " rowspan="& rowspan &">"& sSero1 & firstLine &"</tr>"
				secondLine = "<tr bgcolor='#EEF6FF'>" & secondLine &"</tr>"
	
				response.write firstLine
				response.write secondLine
	
				''-----�����׸� �Ѹ���
				sCOLNM = "SOSOKGB_A"
				sCOLCD = ""
	
				sSQL = " EXEC SP_SUM_BY_HISTORY_BCLASS " & sDepth & ",'" & FromDate & "','" & ToDate & "','TB_LIFECALLHISTORY','"&whereCD2 & "','"& whereCD1 & "','"&sCOLNM&"','Q','','"&sGroupBy&"','"&sSelectCol&"','"&sSelectCol1&"','"&sNullCol&"','"&CHANNELGB&"','"&JEONDOR&"','',''"
	
				'response.write sSQL
				'response.end
				set rsSUM = db.execute(sSQL)
				firstLine = ""
	
				Do Until rsSUM.eof
					sBG = "#ffffff"
					firstLine = ""
					'--------------Ű�� �ش��ϴ� ��
					For i = 1 To sDepth
						sUser = rsSUM("col_"&i)
						'sCodeName = db_GetUserName(sUser)
	
						If IsNull(rsSUM("col_"&i)) Then
							sBG = "#EEF6FF"
							firstLine = firstLine & "<td align='center' class='TDCont' colspan='"&sDepth-i+1&"' WIDTH='200'>��</td>"
							Exit for
						Else
							If i = 1 Then
								sCodeName = db_GetSumColNm(whereCD8,i,rsSUM("col_"&i),"","","","")
							ElseIf i = 2 Then
								sCodeName = db_GetSumColNm(whereCD8,i,rsSUM("col_"&i-1),rsSUM("col_"&i),"","","")
							ElseIf i = 3 Then
								sCodeName = db_GetSumColNm(whereCD8,i,rsSUM("col_"&i-2),rsSUM("col_"&i-1),rsSUM("col_"&i),"","")
							ElseIf i = 4 Then
								sCodeName = db_GetSumColNm(whereCD8,i,rsSUM("col_"&i-3),rsSUM("col_"&i-2),rsSUM("col_"&i-1),rsSUM("col_"&i),"")
							Else
								sCodeName = db_GetSumColNm(whereCD8,i,rsSUM("col_"&i-4),rsSUM("col_"&i-3),rsSUM("col_"&i-2),rsSUM("col_"&i-1),rsSUM("col_"&i))
							End if
							firstLine = firstLine & "<td align='center' class='TDCont' colspan='1' WIDTH='200'>"&sCodeName&"</td>"
						End If
					Next
					'--------------�����׸��� summary
					For i = sDepth + 1 To rsSUM.Fields.count
	
						firstLine = firstLine & "<td align='center' class='TDCont' colspan='1' WIDTH='200'>"&rsSUM("col_"&i)&"</td>"
	
					Next
					firstLine = "<tr bgcolor='"&sBG&"'>" & firstLine
					response.write firstLine & "</tr>"
					rsSUM.movenext
	
					'�Ұ�
				Loop
				%>
	
				</table>
				</div>
				
				<%
	
			elseif whereCD9 = "�δ�3��" then'�δ�	  - 5depth
	
				rowspan = 3
				sSQL = "	select  ACLASS, BCLASS, CCLASS, CLASSNAME from TB_ARMYINFO "
				sSQL = sSQL & "	where ACLASS < 'O' AND BCLASS IS NULL AND CCLASS IS NULL AND DCLASS IS NULL"
				sSQL = sSQL & "	ORDER BY ACLASS, BCLASS, CCLASS "
				set RsCode = db.execute(sSQL)
	
				Do Until rsCode.eof
	
					sCode = RsCode("ACLASS")
					sCodeName = RsCode("CLASSNAME")
	
					'2DEPTH �� ã��
					iCol = 0
					'secondLine = ""
					sCodeList = ""
					sSQL = "	select ACLASS, BCLASS, CCLASS, CLASSNAME from TB_ARMYINFO "
					sSQL = sSQL & "	where ACLASS = '" &sCode&"'  AND BCLASS IS NOT NULL AND CCLASS IS NULL"
					sSQL = sSQL & "	ORDER BY ACLASS, BCLASS, CCLASS "
					set RsCode1 = db.execute(sSQL)
	
					Do Until rsCode1.eof
	
						sCode = RsCode1("BCLASS")
						sCodeName = RsCode1("CLASSNAME")
						iCol = iCol + 1
	
						'3DEPTH �� ã��
						iCol1 = 0
	
						'secondLine = ""
						sSQL = "	select ACLASS, BCLASS, CCLASS, CLASSNAME from TB_ARMYINFO "
						sSQL = sSQL & "	where ACLASS = '" &RsCode("ACLASS")&"'  AND BCLASS = '" & RsCode1("BCLASS")&"' AND CCLASS IS NOT NULL"
						sSQL = sSQL & "	ORDER BY ACLASS, BCLASS, CCLASS, DCLASS"
						set RsCode2 = db.execute(sSQL)
						'-------------------------------------------------------------------------------------------------------------------------------------
						
						Do Until rsCode2.eof
							
							iCol1 = iCol1 + 1
							cols = cols + 1
							iCol = iCol + 1
							sSQL = "INSERT INTO TMP_CODE_VALUE VALUES ("&cols& ",'SOSOKGB_C','" & RsCode2("CCLASS") & "')"
							db.execute(sSQL)
							If sCodeList = "" then
								sCodeList = RsCode2("CCLASS")
							Else
								sCodeList = sCodeList & "|" & RsCode2("CCLASS")
							End if
							
							threeLine = threeLine & "<td align='center' class='TDCont'  width='150'>" & RsCode2("CLASSNAME") & "</td>"
							
							rsCode2.movenext
							'�Ұ�
						Loop
	
						If iCol1 <= 0 Then
							cols = cols + 1
							sSQL = "INSERT INTO TMP_CODE_VALUE VALUES ("&cols& ",'SOSOKGB_B','" & RsCode1("BCLASS") & "')"
							db.execute(sSQL)
							secondLine = secondLine & "<td align='center' class='TDCont' colspan=1 rowspan='2' width='150'>"&RsCode1("CLASSNAME")&"</td>"
						Else
							iCol1 = iCol1 + 1
							cols = cols + 1
							sSQL = "INSERT INTO TMP_CODE_VALUE VALUES ("&cols& ",'SOSOKGB_B','" & RsCode1("BCLASS") & "')"
							db.execute(sSQL)
							'iCol = iCol + 1
							threeLine = threeLine & "<td align='center' class='TDCont'  width='150' rowspan='1'>��</td>"
							secondLine = secondLine & "<td align='center' class='TDCont' colspan="&iCol1&">"&RsCode1("CLASSNAME")&"</td>"
						End If
						iCol1 = 0
						
						rsCode1.movenext
						'�Ұ�
					Loop
	
					If iCol = 0 Then
						cols = cols + 1
						sSQL = "INSERT INTO TMP_CODE_VALUE VALUES ("&cols& ",'SOSOKGB_A','" & RsCode("ACLASS") & "')"
						db.execute(sSQL)
						firstLine = firstLine & "<td align='center' class='TDCont' colspan="&iCol&" rowspan='3' width='150'>"&RsCode("CLASSNAME")&"</td>"
					Else
						cols = cols + 1
						iCol = iCol + 1
						sSQL = "INSERT INTO TMP_CODE_VALUE VALUES ("&cols& ",'SOSOKGB_A','" & RsCode("ACLASS") & "')"
						db.execute(sSQL)
						secondLine = secondLine & "<td align='center' class='TDCont'  width='150' rowspan='2'>��</td>"
						firstLine = firstLine & "<td align='center' class='TDCont' colspan="&iCol&" >"&RsCode("CLASSNAME")&"</td>"
						iCol = 0
					End if
					
					rsCode.movenext
					'�Ұ�
				Loop
				'�Ѱ�
				sWidth = cols * 100
				%>
				
				<DIV style="OVERFLOW-X: auto;OVERFLOW-Y: auto; MARGIN: 0px 0px 0px 0px; WIDTH: 1200; HEIGHT=400;">
				<table width="<%=sWidth%>"  border="1" cellpadding="0" cellspacing="0" bgcolor="#EFECE5" align="center" bordercolor="black" bordercolordark="white" bordercolorlight="black">
					
				<%
				firstLine = firstLine & "<td align='center' class='TDCont' rowspan="& rowspan &" width='150'>��</td>"
				firstLine = "<tr bgcolor='#EEF6FF'>" & sSero & " rowspan="& rowspan &">"& sSero1 & firstLine &"</tr>"
				secondLine = "<tr bgcolor='#EEF6FF'>" & secondLine &"</tr>"
				threeLine = "<tr bgcolor='#EEF6FF'>" & threeLine &"</tr>"
				response.write firstLine
				response.write secondLine
				response.write threeLine
	
				''-----�����׸� �Ѹ���
				sCOLNM = "SOSOKGB_A"
				sCOLCD = ""
	
				sSQL = " EXEC SP_SUM_BY_HISTORY_BCLASS " & sDepth & ",'" & FromDate & "','" & ToDate & "','TB_LIFECALLHISTORY','"&whereCD2 & "','"& whereCD1 & "','"&sCOLNM&"','Q','','"&sGroupBy&"','"&sSelectCol&"','"&sSelectCol1&"','"&sNullCol&"','"&CHANNELGB&"','"&JEONDOR&"','',''"
	
				'response.write sSQL
	
				set rsSUM = db.execute(sSQL)
				firstLine = ""
	
				Do Until rsSUM.eof
					sBG = "#ffffff"
					firstLine = ""
					'--------------Ű�� �ش��ϴ� ��
					For i = 1 To sDepth
						sUser = rsSUM("col_"&i)
						'sCodeName = db_GetUserName(sUser)
	
						If IsNull(rsSUM("col_"&i)) Then
							sBG = "#EEF6FF"
							firstLine = firstLine & "<td align='center' class='TDCont' colspan='"&sDepth-i+1&"' WIDTH='200'>��</td>"
							Exit for
						Else
							If i = 1 Then
								sCodeName = db_GetSumColNm(whereCD8,i,rsSUM("col_"&i),"","","","")
							ElseIf i = 2 Then
								sCodeName = db_GetSumColNm(whereCD8,i,rsSUM("col_"&i-1),rsSUM("col_"&i),"","","")
							ElseIf i = 3 Then
								sCodeName = db_GetSumColNm(whereCD8,i,rsSUM("col_"&i-2),rsSUM("col_"&i-1),rsSUM("col_"&i),"","")
							ElseIf i = 4 Then
								sCodeName = db_GetSumColNm(whereCD8,i,rsSUM("col_"&i-3),rsSUM("col_"&i-2),rsSUM("col_"&i-1),rsSUM("col_"&i),"")
							Else
								sCodeName = db_GetSumColNm(whereCD8,i,rsSUM("col_"&i-4),rsSUM("col_"&i-3),rsSUM("col_"&i-2),rsSUM("col_"&i-1),rsSUM("col_"&i))
							End if
							firstLine = firstLine & "<td align='center' class='TDCont' colspan='1' WIDTH='200'>"&sCodeName&"</td>"
						End If
					Next
					'--------------�����׸��� summary
					For i = sDepth + 1 To rsSUM.Fields.count
	
						firstLine = firstLine & "<td align='center' class='TDCont' colspan='1' WIDTH='200'>"&rsSUM("col_"&i)&"</td>"
	
					Next
					firstLine = "<tr bgcolor='"&sBG&"'>" & firstLine
					response.write firstLine & "</tr>"
					rsSUM.movenext
	
					'�Ұ�
				Loop
				%>
	
				</table>
				</div>
				
				<%
	
			elseif whereCD9 = "���о�" then'���о� - 2depth
	
				rowspan = 2
				sSQL = "	select ACLASS, BCLASS, CCLASS, CLASSNAME from TB_ARMYINFO "
				sSQL = sSQL & "	where ACLASS = 'O' AND BCLASS IS NOT NULL AND CCLASS IS NULL"
				sSQL = sSQL & "	ORDER BY ACLASS, BCLASS, CCLASS "
				set RsCode = db.execute(sSQL)
	
				Do Until rsCode.eof
	
					sCode = RsCode("BCLASS")
					sCodeName = RsCode("CLASSNAME")
	
					'2DEPTH �� ã��
					iCol = 0
					'secondLine = ""
					sCodeList = ""
					sSQL = "	select ACLASS, BCLASS, CCLASS, CLASSNAME from TB_ARMYINFO "
					sSQL = sSQL & "	where ACLASS = 'O' AND BCLASS = '" &sCode&"'  AND CCLASS IS NOT NULL"
					sSQL = sSQL & "	ORDER BY ACLASS, BCLASS, CCLASS "
					set RsCode1 = db.execute(sSQL)
	
					Do Until rsCode1.eof
	
						sCode = RsCode1("CCLASS")
						sCodeName = RsCode1("CLASSNAME")
	
						iCol = iCol + 1
						cols = cols + 1
						sSQL = "INSERT INTO TMP_CODE_VALUE VALUES ("&cols& ",'CALLCLASS_C','" & sCode & "')"
						db.execute(sSQL)
	
						If sCodeList = "" then
							sCodeList = sCode
						Else
							sCodeList = sCodeList & "|" & sCode
						End if
						secondLine = secondLine & "<td align='center' class='TDCont'  width='150'>" & sCodeName & "</td>"
						
						rsCode1.movenext
						'�Ұ�
					Loop
	
					If iCol = 0 Then
						sSQL = "INSERT INTO TMP_CODE_VALUE VALUES ("&cols& ",'CALLCLASS_B','" & sCode & "')"
						db.execute(sSQL)
						cols = cols + 1
						firstLine = firstLine & "<td align='center' class='TDCont' colspan="&iCol+1&" rowspan='2' width='150'>"&RsCode("CLASSNAME")&"</td>"
					Else
						sSQL = "INSERT INTO TMP_CODE_VALUE VALUES ("&cols& ",'CALLCLASS_C','" & sCodeList & "')"
						db.execute(sSQL)
						cols = cols + 1
						secondLine = secondLine & "<td align='center' class='TDCont'  width='150'>��</td>"
	
						firstLine = firstLine & "<td align='center' class='TDCont' colspan="&iCol+1&">"&RsCode("CLASSNAME")&"</td>"
					End if
					
					rsCode.movenext
					'�Ұ�
				Loop
				'�Ѱ�
				sWidth = cols * 150
				%>
				
				<DIV style="OVERFLOW-X: auto;OVERFLOW-Y: auto; MARGIN: 0px 0px 0px 0px; WIDTH: 1200; HEIGHT=400;">
				<table width="<%=sWidth%>"  border="1" cellpadding="0" cellspacing="0" bgcolor="#EFECE5" align="center" bordercolor="black" bordercolordark="white" bordercolorlight="black">
					
				<%
				firstLine = firstLine & "<td align='center' class='TDCont' rowspan="& rowspan &" width='150'>��</td>"
				firstLine = "<tr bgcolor='#EEF6FF'>" & sSero & " rowspan="& rowspan &">"& sSero1 & firstLine &"</tr>"
				secondLine = "<tr bgcolor='#EEF6FF'>" & secondLine &"</tr>"
				response.write firstLine
				response.write secondLine
	
				''-----�����׸� �Ѹ���
				sCOLNM = "CALLCLASS_B"
				sCOLCD = ""
	
				sSQL = " EXEC SP_SUM_BY_HISTORY_BCLASS " & sDepth & ",'" & FromDate & "','" & ToDate & "','TB_LIFECALLHISTORY','"&whereCD2 & "','"& whereCD1 & "','"&sCOLNM&"','Q','','"&sGroupBy&"','"&sSelectCol&"','"&sSelectCol1&"','"&sNullCol&"','"&CHANNELGB&"','"&JEONDOR&"','',''"
	
				'response.write sSQL
	
				set rsSUM = db.execute(sSQL)
				firstLine = ""
	
				Do Until rsSUM.eof
					sBG = "#ffffff"
					firstLine = ""
					'--------------Ű�� �ش��ϴ� ��
					For i = 1 To sDepth
						sUser = rsSUM("col_"&i)
						'sCodeName = db_GetUserName(sUser)
	
						If IsNull(rsSUM("col_"&i)) Then
							sBG = "#EEF6FF"
							firstLine = firstLine & "<td align='center' class='TDCont' colspan='"&sDepth-i+1&"' WIDTH='200'>��</td>"
							Exit for
						Else
							If i = 1 Then
								sCodeName = db_GetSumColNm(whereCD8,i,rsSUM("col_"&i),"","","","")
							ElseIf i = 2 Then
								sCodeName = db_GetSumColNm(whereCD8,i,rsSUM("col_"&i-1),rsSUM("col_"&i),"","","")
							ElseIf i = 3 Then
								sCodeName = db_GetSumColNm(whereCD8,i,rsSUM("col_"&i-2),rsSUM("col_"&i-1),rsSUM("col_"&i),"","")
							ElseIf i = 4 Then
								sCodeName = db_GetSumColNm(whereCD8,i,rsSUM("col_"&i-3),rsSUM("col_"&i-2),rsSUM("col_"&i-1),rsSUM("col_"&i),"")
							Else
								sCodeName = db_GetSumColNm(whereCD8,i,rsSUM("col_"&i-4),rsSUM("col_"&i-3),rsSUM("col_"&i-2),rsSUM("col_"&i-1),rsSUM("col_"&i))
							End if
							firstLine = firstLine & "<td align='center' class='TDCont' colspan='1' WIDTH='200'>"&sCodeName&"</td>"
						End If
					Next
					'--------------�����׸��� summary
					For i = sDepth + 1 To rsSUM.Fields.count
	
						firstLine = firstLine & "<td align='center' class='TDCont' colspan='1' WIDTH='200'>"&rsSUM("col_"&i)&"</td>"
	
					Next
					firstLine = "<tr bgcolor='"&sBG&"'>" & firstLine
					response.write firstLine & "</tr>"
					
					rsSUM.movenext
					'�Ұ�
				Loop
				%>
	
				</table>
				</div>
				
				<%
	
			elseif whereCD9 = "��ġ��" then'��ġ�� - 1depth
	
				rowspan = 2
				sSQL = "	select ACLASS, BCLASS, CCLASS, CLASSNAME from TB_ARMYINFO "
				sSQL = sSQL & "	where ACLASS = 'U' AND BCLASS IS NOT NULL AND CCLASS IS NULL"
				sSQL = sSQL & "	ORDER BY ACLASS, BCLASS, CCLASS "
				set RsCode = db.execute(sSQL)
	
				Do Until rsCode.eof
	
					sCode = RsCode("BCLASS")
					sCodeName = RsCode("CLASSNAME")
	
					'2DEPTH �� ã��
					iCol = 0
					'secondLine = ""
					sCodeList = ""
					sSQL = "	select ACLASS, BCLASS, CCLASS, CLASSNAME from TB_ARMYINFO "
					sSQL = sSQL & "	where ACLASS = 'U' AND BCLASS = '" &sCode&"'  AND CCLASS IS NOT NULL"
					sSQL = sSQL & "	ORDER BY ACLASS, BCLASS, CCLASS "
					set RsCode1 = db.execute(sSQL)
	
					Do Until rsCode1.eof
	
						sCode = RsCode1("CCLASS")
						sCodeName = RsCode1("CLASSNAME")
	
						iCol = iCol + 1
						cols = cols + 1
						sSQL = "INSERT INTO TMP_CODE_VALUE VALUES ("&cols& ",'PROCESSGB_C','" & sCode & "')"
						db.execute(sSQL)
	
						If sCodeList = "" then
							sCodeList = sCode
						Else
							sCodeList = sCodeList & "|" & sCode
						End if
						secondLine = secondLine & "<td align='center' class='TDCont'  width='150'>" & sCodeName & "</td>"
						
						rsCode1.movenext
						'�Ұ�
					Loop
	
					If iCol = 0 Then
						sSQL = "INSERT INTO TMP_CODE_VALUE VALUES ("&cols& ",'PROCESSGB_B','" & sCode & "')"
						db.execute(sSQL)
						cols = cols + 1
						firstLine = firstLine & "<td align='center' class='TDCont' colspan="&iCol+1&" rowspan='2' width='150'>"&RsCode("CLASSNAME")&"</td>"
					Else
						sSQL = "INSERT INTO TMP_CODE_VALUE VALUES ("&cols& ",'PROCESSGB_C','" & sCodeList & "')"
						db.execute(sSQL)
						cols = cols + 1
						secondLine = secondLine & "<td align='center' class='TDCont'  width='150'>��</td>"
	
						firstLine = firstLine & "<td align='center' class='TDCont' colspan="&iCol+1&">"&RsCode("CLASSNAME")&"</td>"
					End if
					
					rsCode.movenext
					'�Ұ�
				Loop
				'�Ѱ�
				sWidth = cols * 150
				%>
				
				<DIV style="OVERFLOW-X: auto;OVERFLOW-Y: auto; MARGIN: 0px 0px 0px 0px; WIDTH: 1200; HEIGHT=400;">
				<table width="<%=sWidth%>"  border="1" cellpadding="0" cellspacing="0" bgcolor="#EFECE5" align="center" bordercolor="black" bordercolordark="white" bordercolorlight="black">
					
				<%
				firstLine = firstLine & "<td align='center' class='TDCont' rowspan="& rowspan &" width='150'>��</td>"
				firstLine = "<tr bgcolor='#EEF6FF'>" & sSero & " rowspan="& rowspan &">"& sSero1 & firstLine &"</tr>"
				secondLine = "<tr bgcolor='#EEF6FF'>" & secondLine &"</tr>"
				response.write firstLine
				response.write secondLine
	
				''-----�����׸� �Ѹ���
				sCOLNM = "PROCESSGB_B"
				sCOLCD = ""
	
				sSQL = " EXEC SP_SUM_BY_HISTORY_BCLASS " & sDepth & ",'" & FromDate & "','" & ToDate & "','TB_LIFECALLHISTORY','"&whereCD2 & "','"& whereCD1 & "','"&sCOLNM&"','U','','"&sGroupBy&"','"&sSelectCol&"','"&sSelectCol1&"','"&sNullCol&"','"&CHANNELGB&"','"&JEONDOR&"','',''"
	
				'response.write sSQL
	
				set rsSUM = db.execute(sSQL)
				firstLine = ""
	
				Do Until rsSUM.eof
					sBG = "#ffffff"
					firstLine = ""
					'--------------Ű�� �ش��ϴ� ��
					For i = 1 To sDepth
						sUser = rsSUM("col_"&i)
						'sCodeName = db_GetUserName(sUser)
	
						If IsNull(rsSUM("col_"&i)) Then
							sBG = "#EEF6FF"
							firstLine = firstLine & "<td align='center' class='TDCont' colspan='"&sDepth-i+1&"' WIDTH='200'>��</td>"
							Exit for
						Else
							If i = 1 Then
								sCodeName = db_GetSumColNm(whereCD8,i,rsSUM("col_"&i),"","","","")
							ElseIf i = 2 Then
								sCodeName = db_GetSumColNm(whereCD8,i,rsSUM("col_"&i-1),rsSUM("col_"&i),"","","")
							ElseIf i = 3 Then
								sCodeName = db_GetSumColNm(whereCD8,i,rsSUM("col_"&i-2),rsSUM("col_"&i-1),rsSUM("col_"&i),"","")
							ElseIf i = 4 Then
								sCodeName = db_GetSumColNm(whereCD8,i,rsSUM("col_"&i-3),rsSUM("col_"&i-2),rsSUM("col_"&i-1),rsSUM("col_"&i),"")
							Else
								sCodeName = db_GetSumColNm(whereCD8,i,rsSUM("col_"&i-4),rsSUM("col_"&i-3),rsSUM("col_"&i-2),rsSUM("col_"&i-1),rsSUM("col_"&i))
							End if
							firstLine = firstLine & "<td align='center' class='TDCont' colspan='1' WIDTH='200'>"&sCodeName&"</td>"
						End If
					Next
					'--------------�����׸��� summary
					For i = sDepth + 1 To rsSUM.Fields.count
	
						firstLine = firstLine & "<td align='center' class='TDCont' colspan='1' WIDTH='200'>"&rsSUM("col_"&i)&"</td>"
	
					Next
					firstLine = "<tr bgcolor='"&sBG&"'>" & firstLine
					response.write firstLine & "</tr>"
					
					rsSUM.movenext
					'�Ұ�
				Loop
				%>
	
				</table>
				</div>
				
				<%
	
			elseif whereCD9 = "������" then'���� - 1depth
				
				rowspan = 1
				sSQL = "	select CODE, CODENAME from TB_CODE "
				sSQL = sSQL & "	where CODEGROUP = 'C11' AND USEYN = 'Y'"
				sSQL = sSQL & "	ORDER BY CODE "
				set RsCode = db.execute(sSQL)
	
				Do Until rsCode.eof
	
					sCode = RsCode("CODE")
					sCodeName = RsCode("CODENAME")
	
					'2DEPTH �� ã��
					iCol = iCol + 1
					cols = cols + 1
	
					firstLine = firstLine & "<td align='center' class='TDCont' colspan='1' WIDTH='200'>"&RsCode("CODENAME")&"</td>"
					
					rsCode.movenext
					'�Ұ�
				Loop
				cols = cols + 1
				firstLine = firstLine & "<td align='center' class='TDCont' colspan='1' WIDTH='200'>��</td>"
				'�Ѱ�
				sWidth = cols * 200
				%>
				
				<DIV style="OVERFLOW-X: auto;OVERFLOW-Y: auto; MARGIN: 0px 0px 0px 0px; WIDTH: 1200; HEIGHT='400';">
				<table width="<%=sWidth%>"  border="1" cellpadding="0" cellspacing="0" bgcolor="#EFECE5" align="center" bordercolor="black" bordercolordark="white" bordercolorlight="black">
					
				<%
				firstLine = "<tr bgcolor='#EEF6FF'>" & sSero & " rowspan="& rowspan &">"& sSero1 & firstLine &"</tr>"
				response.write firstLine
	
				sCOLNM = "WEATHER"
				sCOLCD = "C11"
	
				sSQL = " EXEC SP_SUM_BY_CODE " & sDepth & ",'" & FromDate & "','" & ToDate & "','TB_LIFECALLHISTORY','"&whereCD2 & "','"& whereCD1 & "','"&sCOLNM&"','"&sCOLCD&"','"&sGroupBy&"','"&sSelectCol&"','"&sSelectCol1&"','"&sNullCol&"','"&CHANNELGB&"','"&JEONDOR&"','',''"
	
				'response.write sSQL
	
				set rsSUM = db.execute(sSQL)
				firstLine = ""
	
				Do Until rsSUM.eof
					sBG = "#ffffff"
					firstLine = ""
					For i = 1 To sDepth
						sUser = rsSUM("col_"&i)
						If IsNull(rsSUM("col_"&i)) Then
							sBG = "#EEF6FF"
							firstLine = firstLine & "<td align='center' class='TDCont' colspan='"&sDepth-i+1&"' WIDTH='200'>��</td>"
							Exit for
						Else
							If i = 1 Then
								sCodeName = db_GetSumColNm(whereCD8,i,rsSUM("col_"&i),"","","","")
							ElseIf i = 2 Then
								sCodeName = db_GetSumColNm(whereCD8,i,rsSUM("col_"&i-1),rsSUM("col_"&i),"","","")
							ElseIf i = 3 Then
								sCodeName = db_GetSumColNm(whereCD8,i,rsSUM("col_"&i-2),rsSUM("col_"&i-1),rsSUM("col_"&i),"","")
							ElseIf i = 4 Then
								sCodeName = db_GetSumColNm(whereCD8,i,rsSUM("col_"&i-3),rsSUM("col_"&i-2),rsSUM("col_"&i-1),rsSUM("col_"&i),"")
							Else
								sCodeName = db_GetSumColNm(whereCD8,i,rsSUM("col_"&i-4),rsSUM("col_"&i-3),rsSUM("col_"&i-2),rsSUM("col_"&i-1),rsSUM("col_"&i))
							End if
	
							firstLine = firstLine & "<td align='center' class='TDCont' colspan='1' WIDTH='200'>"&sCodeName&"</td>"
						End If
					Next
	
					For i = sDepth + 1 To rsSUM.Fields.count
	
						firstLine = firstLine & "<td align='center' class='TDCont' colspan='1' WIDTH='200'>"&rsSUM("col_"&i)&"</td>"
	
					Next
					firstLine = "<tr bgcolor='"&sBG&"'>" & firstLine
					response.write firstLine & "</tr>"
					
					rsSUM.movenext
					'�Ұ�
				Loop
				%>
	
				</table>
				</div>
				
				<%
			elseif whereCD9 = "�������" then'������� - 1depth
				
				rowspan = 1
				sSQL = "	select CODE, CODENAME from TB_CODE "
				sSQL = sSQL & "	where CODEGROUP = 'C10' AND USEYN = 'Y'"
				sSQL = sSQL & "	ORDER BY CODE "
				set RsCode = db.execute(sSQL)
	
				Do Until rsCode.eof
	
					sCode = RsCode("CODE")
					sCodeName = RsCode("CODENAME")
	
					'2DEPTH �� ã��
					iCol = iCol + 1
					cols = cols + 1
	
					firstLine = firstLine & "<td align='center' class='TDCont' colspan='1' WIDTH='150'>"&RsCode("CODENAME")&"</td>"
					
					rsCode.movenext
					'�Ұ�
				Loop
				cols = cols + 1
				firstLine = firstLine & "<td align='center' class='TDCont' colspan='1' WIDTH='150'>��</td>"
				'�Ѱ�
				sWidth = cols * 150
				%>
				
				<DIV style="OVERFLOW-X: auto;OVERFLOW-Y: auto; MARGIN: 0px 0px 0px 0px; WIDTH: 1200; HEIGHT='400';">
				<table width="<%=sWidth%>"  border="1" cellpadding="0" cellspacing="0" bgcolor="#EFECE5" align="center" bordercolor="black" bordercolordark="white" bordercolorlight="black">
					
				<%
				firstLine = "<tr bgcolor='#EEF6FF'>" & sSero & " rowspan="& rowspan &">"& sSero1 & firstLine &"</tr>"
				response.write firstLine
	
				sCOLNM = "CALLFLAG"
				sCOLCD = "C10"
	
				sSQL = " EXEC SP_SUM_BY_CODE " & sDepth & ",'" & FromDate & "','" & ToDate & "','TB_LIFECALLHISTORY','"&whereCD2 & "','"& whereCD1 & "','"&sCOLNM&"','"&sCOLCD&"','"&sGroupBy&"','"&sSelectCol&"','"&sSelectCol1&"','"&sNullCol&"','"&CHANNELGB&"','"&JEONDOR&"','',''"
	
				'response.write sSQL
	
				set rsSUM = db.execute(sSQL)
				firstLine = ""
	
				Do Until rsSUM.eof
					sBG = "#ffffff"
					firstLine = ""
					For i = 1 To sDepth
						sUser = rsSUM("col_"&i)
						'sCodeName = db_GetUserName(sUser)
	
						If IsNull(rsSUM("col_"&i)) Then
							sBG = "#EEF6FF"
							firstLine = firstLine & "<td align='center' class='TDCont' colspan='"&sDepth-i+1&"' WIDTH='200'>��</td>"
							Exit for
						Else
							If i = 1 Then
								sCodeName = db_GetSumColNm(whereCD8,i,rsSUM("col_"&i),"","","","")
							ElseIf i = 2 Then
								sCodeName = db_GetSumColNm(whereCD8,i,rsSUM("col_"&i-1),rsSUM("col_"&i),"","","")
							ElseIf i = 3 Then
								sCodeName = db_GetSumColNm(whereCD8,i,rsSUM("col_"&i-2),rsSUM("col_"&i-1),rsSUM("col_"&i),"","")
							ElseIf i = 4 Then
								sCodeName = db_GetSumColNm(whereCD8,i,rsSUM("col_"&i-3),rsSUM("col_"&i-2),rsSUM("col_"&i-1),rsSUM("col_"&i),"")
							Else
								sCodeName = db_GetSumColNm(whereCD8,i,rsSUM("col_"&i-4),rsSUM("col_"&i-3),rsSUM("col_"&i-2),rsSUM("col_"&i-1),rsSUM("col_"&i))
							End if
							firstLine = firstLine & "<td align='center' class='TDCont' colspan='1' WIDTH='200'>"&sCodeName&"</td>"
						End If
					Next
	
					For i = sDepth + 1 To rsSUM.Fields.count
	
						firstLine = firstLine & "<td align='center' class='TDCont' colspan='1' WIDTH='200'>"&rsSUM("col_"&i)&"</td>"
	
					Next
					firstLine = "<tr bgcolor='"&sBG&"'>" & firstLine
					response.write firstLine & "</tr>"
					
					rsSUM.movenext
					'�Ұ�
				Loop
				%>
	
				</table>
				</div>
				
				<%
			elseif whereCD9 = "��������" then'�������� - 1depth
				
				rowspan = 1
				sSQL = "	select CODE, CODENAME from TB_CODE "
				sSQL = sSQL & "	where CODEGROUP = 'C12' AND USEYN = 'Y'"
				sSQL = sSQL & "	ORDER BY CODE "
				set RsCode = db.execute(sSQL)
	
				Do Until rsCode.eof
	
					sCode = RsCode("CODE")
					sCodeName = RsCode("CODENAME")
	
					'2DEPTH �� ã��
					iCol = iCol + 1
					cols = cols + 1
	
					firstLine = firstLine & "<td align='center' class='TDCont' colspan='1' WIDTH='150'>"&RsCode("CODENAME")&"</td>"
					
					rsCode.movenext
					'�Ұ�
				Loop
				cols = cols + 1
				firstLine = firstLine & "<td align='center' class='TDCont' colspan='1' WIDTH='150'>��</td>"
				'�Ѱ�
				sWidth = cols * 150
				%>
				
				<DIV style="OVERFLOW-X: auto;OVERFLOW-Y: auto; MARGIN: 0px 0px 0px 0px; WIDTH: 1200; HEIGHT='400';">
				<table width="<%=sWidth%>"  border="1" cellpadding="0" cellspacing="0" bgcolor="#EFECE5" align="center" bordercolor="black" bordercolordark="white" bordercolorlight="black">
					
				<%
				firstLine = "<tr bgcolor='#EEF6FF'>" & sSero & " rowspan="& rowspan &">"& sSero1 & firstLine &"</tr>"
				response.write firstLine
	
				sCOLNM = "FAMILYGB"
				sCOLCD = "C12"
	
				sSQL = " EXEC SP_SUM_BY_CODE " & sDepth & ",'" & FromDate & "','" & ToDate & "','TB_LIFECALLHISTORY','"&whereCD2 & "','"& whereCD1 & "','"&sCOLNM&"','"&sCOLCD&"','"&sGroupBy&"','"&sSelectCol&"','"&sSelectCol1&"','"&sNullCol&"','"&CHANNELGB&"','"&JEONDOR&"','',''"
	
				'response.write sSQL
	
				set rsSUM = db.execute(sSQL)
				firstLine = ""
	
				Do Until rsSUM.eof
					sBG = "#ffffff"
					firstLine = ""
					For i = 1 To sDepth
						sUser = rsSUM("col_"&i)
						'sCodeName = db_GetUserName(sUser)
	
						If IsNull(rsSUM("col_"&i)) Then
							sBG = "#EEF6FF"
							firstLine = firstLine & "<td align='center' class='TDCont' colspan='"&sDepth-i+1&"' WIDTH='200'>��</td>"
							Exit for
						Else
							If i = 1 Then
								sCodeName = db_GetSumColNm(whereCD8,i,rsSUM("col_"&i),"","","","")
							ElseIf i = 2 Then
								sCodeName = db_GetSumColNm(whereCD8,i,rsSUM("col_"&i-1),rsSUM("col_"&i),"","","")
							ElseIf i = 3 Then
								sCodeName = db_GetSumColNm(whereCD8,i,rsSUM("col_"&i-2),rsSUM("col_"&i-1),rsSUM("col_"&i),"","")
							ElseIf i = 4 Then
								sCodeName = db_GetSumColNm(whereCD8,i,rsSUM("col_"&i-3),rsSUM("col_"&i-2),rsSUM("col_"&i-1),rsSUM("col_"&i),"")
							Else
								sCodeName = db_GetSumColNm(whereCD8,i,rsSUM("col_"&i-4),rsSUM("col_"&i-3),rsSUM("col_"&i-2),rsSUM("col_"&i-1),rsSUM("col_"&i))
							End if
							firstLine = firstLine & "<td align='center' class='TDCont' colspan='1' WIDTH='200'>"&sCodeName&"</td>"
						End If
					Next
	
					For i = sDepth + 1 To rsSUM.Fields.count
	
						firstLine = firstLine & "<td align='center' class='TDCont' colspan='1' WIDTH='200'>"&rsSUM("col_"&i)&"</td>"
	
					Next
					firstLine = "<tr bgcolor='"&sBG&"'>" & firstLine
					response.write firstLine & "</tr>"
					
					rsSUM.movenext
					'�Ұ�
				Loop
	
				%>
	
				</table>
				</div>
				
				<%
			elseif whereCD9 = "����������" then'���������� - 1depth
				
				rowspan = 1
				sSQL = "	select ACLASS, BCLASS, CCLASS, CLASSNAME from TB_ARMYINFO "
				sSQL = sSQL & "	where ACLASS = 'R' AND BCLASS IS NOT NULL AND CCLASS IS NULL"
				sSQL = sSQL & "	ORDER BY ACLASS, BCLASS, CCLASS "
				set RsCode = db.execute(sSQL)
	
				Do Until rsCode.eof
	
					sCode = RsCode("BCLASS")
					sCodeName = RsCode("CLASSNAME")
	
					cols = cols + 1
	
					firstLine = firstLine & "<td align='center' class='TDCont' colspan=1 width='150'>"&RsCode("CLASSNAME")&"</td>"
	
					rsCode.movenext
					'�Ұ�
				Loop
				'�Ѱ�
				cols = cols + 1
				firstLine = firstLine & "<td align='center' class='TDCont' colspan=1 width='150'>��</td>"
				sWidth = cols * 100
				%>
				
				<DIV style="OVERFLOW-X: auto;OVERFLOW-Y: auto; MARGIN: 0px 0px 0px 0px; WIDTH: 1200; HEIGHT=400;">
				<table width="<%=sWidth%>"  border="1" cellpadding="0" cellspacing="0" bgcolor="#EFECE5" align="center" bordercolor="black" bordercolordark="white" bordercolorlight="black">
					
				<%
	
				firstLine = "<tr bgcolor='#EEF6FF'>" & sSero & " rowspan="& rowspan &">"& sSero1 & firstLine &"</tr>"
				response.write firstLine
	
				'---- ����������
				sCOLNM = "CALLKIND_B"
				sCOLCD = "R"
	
				sSQL = " EXEC SP_SUM_BY_HISTORY_ACLASS " & sDepth & ",'" & FromDate & "','" & ToDate & "','TB_LIFECALLHISTORY','"&whereCD2 & "','"& whereCD1 & "','"&sCOLNM&"','"&sCOLCD&"','"&sGroupBy&"','"&sSelectCol&"','"&sSelectCol1&"','"&sNullCol&"','"&CHANNELGB&"','"&JEONDOR&"','',''"
	
				'response.write sSQL
	
				set rsSUM = db.execute(sSQL)
				firstLine = ""

				Do Until rsSUM.eof
					sBG = "#ffffff"
					firstLine = ""
					For i = 1 To sDepth
						sUser = rsSUM("col_"&i)
						'sCodeName = db_GetUserName(sUser)
	
						If IsNull(rsSUM("col_"&i)) Then
							sBG = "#EEF6FF"
							firstLine = firstLine & "<td align='center' class='TDCont' colspan='"&sDepth-i+1&"' WIDTH='200'>��</td>"
							Exit for
						Else
							If i = 1 Then
								sCodeName = db_GetSumColNm(whereCD8,i,rsSUM("col_"&i),"","","","")
							ElseIf i = 2 Then
								sCodeName = db_GetSumColNm(whereCD8,i,rsSUM("col_"&i-1),rsSUM("col_"&i),"","","")
							ElseIf i = 3 Then
								sCodeName = db_GetSumColNm(whereCD8,i,rsSUM("col_"&i-2),rsSUM("col_"&i-1),rsSUM("col_"&i),"","")
							ElseIf i = 4 Then
								sCodeName = db_GetSumColNm(whereCD8,i,rsSUM("col_"&i-3),rsSUM("col_"&i-2),rsSUM("col_"&i-1),rsSUM("col_"&i),"")
							Else
								sCodeName = db_GetSumColNm(whereCD8,i,rsSUM("col_"&i-4),rsSUM("col_"&i-3),rsSUM("col_"&i-2),rsSUM("col_"&i-1),rsSUM("col_"&i))
							End if
							firstLine = firstLine & "<td align='center' class='TDCont' colspan='1' WIDTH='200'>"&sCodeName&"</td>"
						End If
					Next
	
					For i = sDepth + 1 To rsSUM.Fields.count
	
						firstLine = firstLine & "<td align='center' class='TDCont' colspan='1' WIDTH='200'>"&rsSUM("col_"&i)&"</td>"
	
					Next
					firstLine = "<tr bgcolor='"&sBG&"'>" & firstLine
					response.write firstLine & "</tr>"
					
					rsSUM.movenext
					'�Ұ�
				Loop
				%>
	
				</table>
				</div>
				
				<%
			elseif whereCD9 = "����" then'���� - 1depth
	
				rowspan = 1
				sSQL = "	select distinct INCODE FROM TB_LIFECALLHISTORY"
				sSQL = sSQL & "	where JUBDATE >= '" & FROMDATE &"'"
				sSQL = sSQL & "	AND JUBDATE <= '" & TODATE &"'"
				if len(CHANNELGB1) > 0 or len(CHANNELGB2) > 0 or len(CHANNELGB3) > 0 or len(CHANNELGB4) > 0 then
					sSQL = sSQL & " and CHANNELGB in ('" & CHANNELGB1 & "','" & CHANNELGB2 & "','" & CHANNELGB3 & "','" & CHANNELGB4 & "') "
				end if
				set RsCode1 = db.execute(sSQL)
	
				Do Until rsCode1.eof
	
					sCode = RsCode1("INCODE")
					sCodeName= db_getUserName(RsCode1("INCODE"))
	
					'2DEPTH �� ã��
					iCol = iCol + 1
					cols = cols + 1
	
					firstLine = firstLine & "<td align='center' class='TDCont' colspan='1' WIDTH='200'>"& sCodeName&"</td>"
					rsCode1.movenext
	
					'�Ұ�
				Loop
				cols = cols + 1
				firstLine = firstLine & "<td align='center' class='TDCont' colspan='1' WIDTH='200'>��</td>"
				'�Ѱ�
				sWidth = cols * 200
				%>
				
				<DIV style="OVERFLOW-X: auto;OVERFLOW-Y: auto; MARGIN: 0px 0px 0px 0px; WIDTH: 1200; HEIGHT='400';">
				<table width="<%=sWidth%>"  border="1" cellpadding="0" cellspacing="0" bgcolor="#EFECE5" align="center" bordercolor="black" bordercolordark="white" bordercolorlight="black">
					
				<%
				firstLine = "<tr bgcolor='#EEF6FF'>" & sSero & " rowspan="& rowspan &">"& sSero1 & firstLine &"</tr>"
				response.write firstLine
	
				sCOLNM = "INCODE"
				sCOLCD = ""
	
				sSQL = " EXEC SP_SUM_BY_INCODE " & sDepth & ",'" & FromDate & "','" & ToDate & "','TB_LIFECALLHISTORY','"&whereCD2 & "','"& whereCD1 & "','"&sCOLNM&"','"&sCOLCD&"','"&sGroupBy&"','"&sSelectCol&"','"&sSelectCol1&"','"&sNullCol&"','"&CHANNELGB&"','"&JEONDOR&"','',''"
				'response.write sSQL
	
				set rsSUM = db.execute(sSQL)
				firstLine = ""
	
				Do Until rsSUM.eof
					sBG = "#ffffff"
					firstLine = ""
					For i = 1 To sDepth
						sUser = rsSUM("col_"&i)
						'sCodeName = db_GetUserName(sUser)
	
						If IsNull(rsSUM("col_"&i)) Then
							sBG = "#EEF6FF"
							firstLine = firstLine & "<td align='center' class='TDCont' colspan='"&sDepth-i+1&"' WIDTH='200'>��</td>"
							Exit for
						Else
							If i = 1 Then
								sCodeName = db_GetSumColNm(whereCD8,i,rsSUM("col_"&i),"","","","")
							ElseIf i = 2 Then
								sCodeName = db_GetSumColNm(whereCD8,i,rsSUM("col_"&i-1),rsSUM("col_"&i),"","","")
							ElseIf i = 3 Then
								sCodeName = db_GetSumColNm(whereCD8,i,rsSUM("col_"&i-2),rsSUM("col_"&i-1),rsSUM("col_"&i),"","")
							ElseIf i = 4 Then
								sCodeName = db_GetSumColNm(whereCD8,i,rsSUM("col_"&i-3),rsSUM("col_"&i-2),rsSUM("col_"&i-1),rsSUM("col_"&i),"")
							Else
								sCodeName = db_GetSumColNm(whereCD8,i,rsSUM("col_"&i-4),rsSUM("col_"&i-3),rsSUM("col_"&i-2),rsSUM("col_"&i-1),rsSUM("col_"&i))
							End if
							firstLine = firstLine & "<td align='center' class='TDCont' colspan='1' WIDTH='200'>"&sCodeName&"</td>"
						End If
					Next
	
					For i = sDepth + 1 To rsSUM.Fields.count
	
						firstLine = firstLine & "<td align='center' class='TDCont' colspan='1' WIDTH='200'>"&rsSUM("col_"&i)&"</td>"
	
					Next
					firstLine = "<tr bgcolor='"&sBG&"'>" & firstLine
	
					firstLine = firstLine & "</tr>"
					response.write firstLine
					
					rsSUM.movenext
					'�Ұ�
				Loop
				%>
	
				</table>
				</div>
				
				<%
	
			elseif whereCD9 = "�ð�" then'�ð� - 1depth
	
				rowspan = 1
	
				For i = 0 To 23
	
					sCode = i
					sCodeName  = i & "��"
	
					'2DEPTH �� ã��
					iCol = iCol + 1
					cols = cols + 1
	
					firstLine = firstLine & "<td align='center' class='TDCont' colspan='1' WIDTH='120'>"& sCodeName&"</td>"
				Next
				cols = cols + 1
				firstLine = firstLine & "<td align='center' class='TDCont' colspan='1' WIDTH='120'>��</td>"
				sWidth = cols * 120
				%>
				
				<DIV style="OVERFLOW-X: auto;OVERFLOW-Y: auto; MARGIN: 0px 0px 0px 0px; WIDTH: 1200; HEIGHT='400';">
				<table width="<%=sWidth%>"  border="1" cellpadding="0" cellspacing="0" bgcolor="#EFECE5" align="center" bordercolor="black" bordercolordark="white" bordercolorlight="black">
					
				<%
				firstLine = "<tr bgcolor='#EEF6FF'>" & sSero & " rowspan="& rowspan &">"& sSero1 & firstLine &"</tr>"
				response.write firstLine
	
				'sGroupBy = "datepart(hour,JUBTIME)"
	
				'---- �Ҽ�1��
				sCOLNM = "datepart(hour,JUBTIME)"
				sCOLCD = ""

				sSQL = " EXEC SP_SUM_BY_HISTORY_HOUR " & sDepth & ",'" & FromDate & "','" & ToDate & "','TB_LIFECALLHISTORY','"&whereCD2 & "','"& whereCD1 & "','"&sCOLNM&"','"&sCOLCD&"','"&sGroupBy&"','"&sSelectCol&"','"&sSelectCol1&"','"&sNullCol&"','"&CHANNELGB&"','"&JEONDOR&"','',''"
	
				'response.write sSQL
	
				set rsSUM = db.execute(sSQL)
				firstLine = ""
	
				Do Until rsSUM.eof
					sBG = "#ffffff"
					firstLine = ""
					For i = 1 To sDepth
						sUser = rsSUM("col_"&i)
						'sCodeName = db_GetUserName(sUser)
	
						If IsNull(rsSUM("col_"&i)) Then
							sBG = "#EEF6FF"
							firstLine = firstLine & "<td align='center' class='TDCont' colspan='"&sDepth-i+1&"' WIDTH='120'>��</td>"
							Exit for
						Else
							If i = 1 Then
								sCodeName = db_GetSumColNm(whereCD8,i,rsSUM("col_"&i),"","","","")
							ElseIf i = 2 Then
								sCodeName = db_GetSumColNm(whereCD8,i,rsSUM("col_"&i-1),rsSUM("col_"&i),"","","")
							ElseIf i = 3 Then
								sCodeName = db_GetSumColNm(whereCD8,i,rsSUM("col_"&i-2),rsSUM("col_"&i-1),rsSUM("col_"&i),"","")
							ElseIf i = 4 Then
								sCodeName = db_GetSumColNm(whereCD8,i,rsSUM("col_"&i-3),rsSUM("col_"&i-2),rsSUM("col_"&i-1),rsSUM("col_"&i),"")
							Else
								sCodeName = db_GetSumColNm(whereCD8,i,rsSUM("col_"&i-4),rsSUM("col_"&i-3),rsSUM("col_"&i-2),rsSUM("col_"&i-1),rsSUM("col_"&i))
							End if
							firstLine = firstLine & "<td align='center' class='TDCont' colspan='1' WIDTH='120'>"&sCodeName&"</td>"
						End If
					Next
	
					For i = sDepth + 1 To rsSUM.Fields.count
	
						firstLine = firstLine & "<td align='center' class='TDCont' colspan='1' WIDTH='120'>"&rsSUM("col_"&i)&"</td>"
	
					Next
					firstLine = "<tr bgcolor='"&sBG&"'>" & firstLine
					response.write firstLine & "</tr>"
					
					rsSUM.movenext
					'�Ұ�
				Loop
				%>
	
				</table>
				</div>
				
				<%
	
			elseif whereCD9 = "����" then'���� - 1depth
				
				rowspan = 1
	
				firstLine = firstLine & "<td align='center' class='TDCont' colspan='1' WIDTH='150'>��</td>"
				firstLine = firstLine & "<td align='center' class='TDCont' colspan='1' WIDTH='150'>��</td>"
				firstLine = firstLine & "<td align='center' class='TDCont' colspan='1' WIDTH='150'>ȭ</td>"
				firstLine = firstLine & "<td align='center' class='TDCont' colspan='1' WIDTH='150'>��</td>"
				firstLine = firstLine & "<td align='center' class='TDCont' colspan='1' WIDTH='150'>��</td>"
				firstLine = firstLine & "<td align='center' class='TDCont' colspan='1' WIDTH='150'>��</td>"
				firstLine = firstLine & "<td align='center' class='TDCont' colspan='1' WIDTH='150'>��</td>"
				cols = cols + 8
				firstLine = firstLine & "<td align='center' class='TDCont' colspan='1' WIDTH='150'>��</td>"
				sWidth = cols * 150
				%>
				
				<DIV style="OVERFLOW-X: auto;OVERFLOW-Y: auto; MARGIN: 0px 0px 0px 0px; WIDTH: 1200; HEIGHT='400';">
				<table width="<%=sWidth%>"  border="1" cellpadding="0" cellspacing="0" bgcolor="#EFECE5" align="center" bordercolor="black" bordercolordark="white" bordercolorlight="black">
					
				<%
				firstLine = "<tr bgcolor='#EEF6FF'>" & sSero & " rowspan="& rowspan &">"& sSero1 & firstLine &"</tr>"
				response.write firstLine
	
				'---- �Ҽ�1��
				sCOLNM = "datepart(WEEKDAY,JUBTIME)"
				sCOLCD = ""
	
				sSQL = " EXEC SP_SUM_BY_HISTORY_WEEKDAY " & sDepth & ",'" & FromDate & "','" & ToDate & "','TB_LIFECALLHISTORY','"&whereCD2 & "','"& whereCD1 & "','"&sCOLNM&"','"&sCOLCD&"','"&sGroupBy&"','"&sSelectCol&"','"&sSelectCol1&"','"&sNullCol&"','"&CHANNELGB&"','"&JEONDOR&"','',''"
	
				'response.write sSQL
	
				set rsSUM = db.execute(sSQL)
				firstLine = ""
	
				Do Until rsSUM.eof
					sBG = "#ffffff"
					firstLine = ""
					For i = 1 To sDepth
						sUser = rsSUM("col_"&i)
						'sCodeName = db_GetUserName(sUser)
	
						If IsNull(rsSUM("col_"&i)) Then
							sBG = "#EEF6FF"
							firstLine = firstLine & "<td align='center' class='TDCont' colspan='"&sDepth-i+1&"' WIDTH='200'>��</td>"
							Exit for
						Else
							If i = 1 Then
								sCodeName = db_GetSumColNm(whereCD8,i,rsSUM("col_"&i),"","","","")
							ElseIf i = 2 Then
								sCodeName = db_GetSumColNm(whereCD8,i,rsSUM("col_"&i-1),rsSUM("col_"&i),"","","")
							ElseIf i = 3 Then
								sCodeName = db_GetSumColNm(whereCD8,i,rsSUM("col_"&i-2),rsSUM("col_"&i-1),rsSUM("col_"&i),"","")
							ElseIf i = 4 Then
								sCodeName = db_GetSumColNm(whereCD8,i,rsSUM("col_"&i-3),rsSUM("col_"&i-2),rsSUM("col_"&i-1),rsSUM("col_"&i),"")
							Else
								sCodeName = db_GetSumColNm(whereCD8,i,rsSUM("col_"&i-4),rsSUM("col_"&i-3),rsSUM("col_"&i-2),rsSUM("col_"&i-1),rsSUM("col_"&i))
							End if
							firstLine = firstLine & "<td align='center' class='TDCont' colspan='1' WIDTH='200'>"&sCodeName&"</td>"
						End If
					Next
	
					For i = sDepth + 1 To rsSUM.Fields.count
	
						firstLine = firstLine & "<td align='center' class='TDCont' colspan='1' WIDTH='200'>"&rsSUM("col_"&i)&"</td>"
	
					Next
					firstLine = "<tr bgcolor='"&sBG&"'>" & firstLine
					response.write firstLine & "</tr>"
					
					rsSUM.movenext
					'�Ұ�
				Loop
	
				%>
	
				</table>
				</div>
				
				<%
			elseif whereCD9 = "��ȭ�ð�" then'��ȭ�ð� - 1depth
	
				rowspan = 1
				firstLine = firstLine & "<td align='center' class='TDCont' width='150'>1�й̸�</td>"
				firstLine = firstLine & "<td align='center' class='TDCont' width='150'>1-5��</td>"
				firstLine = firstLine & "<td align='center' class='TDCont' width='150'>6-10��</td>"
				firstLine = firstLine & "<td align='center' class='TDCont' width='150'>11-20��</td>"
				firstLine = firstLine & "<td align='center' class='TDCont' width='150'>21-30��</td>"
				firstLine = firstLine & "<td align='center' class='TDCont' width='150'>31-40��</td>"
				firstLine = firstLine & "<td align='center' class='TDCont' width='150'>41-50��</td>"
				firstLine = firstLine & "<td align='center' class='TDCont' width='150'>51-60��</td>"
				firstLine = firstLine & "<td align='center' class='TDCont' width='150'>60���̻�</td>"
				firstLine = firstLine & "<td align='center' class='TDCont' width='150'>��</td>"
				cols = cols + 10
				sWidth = cols * 120
				%>
				
				<DIV style="OVERFLOW-X: auto;OVERFLOW-Y: auto; MARGIN: 0px 0px 0px 0px; WIDTH: 1200; HEIGHT='400';">
				<table width="<%=sWidth%>"  border="1" cellpadding="0" cellspacing="0" bgcolor="#EFECE5" align="center" bordercolor="black" bordercolordark="white" bordercolorlight="black">
					
				<%
				firstLine = "<tr bgcolor='#EEF6FF'>" & sSero & " rowspan="& rowspan &">"& sSero1 & firstLine &"</tr>"
				response.write firstLine
	
				'---- �Ҽ�1��
				sCOLNM = "CALLTIME"
				sCOLCD = ""

				sSQL = " EXEC SP_SUM_BY_HISTORY_CALLTIME " & sDepth & ",'" & FromDate & "','" & ToDate & "','TB_LIFECALLHISTORY','"&whereCD2 & "','"& whereCD1 & "','"&sCOLNM&"','"&sCOLCD&"','"&sGroupBy&"','"&sSelectCol&"','"&sSelectCol1&"','"&sNullCol&"','"&CHANNELGB&"','"&JEONDOR&"','',''"
	
				'response.write sSQL
	
				set rsSUM = db.execute(sSQL)
				firstLine = ""
	
				Do Until rsSUM.eof
					sBG = "#ffffff"
					firstLine = ""
					For i = 1 To sDepth
						sUser = rsSUM("col_"&i)
						'sCodeName = db_GetUserName(sUser)
	
						If IsNull(rsSUM("col_"&i)) Then
							sBG = "#EEF6FF"
							firstLine = firstLine & "<td align='center' class='TDCont' colspan='"&sDepth-i+1&"' WIDTH='200'>��</td>"
							Exit for
						Else
							If i = 1 Then
								sCodeName = db_GetSumColNm(whereCD8,i,rsSUM("col_"&i),"","","","")
							ElseIf i = 2 Then
								sCodeName = db_GetSumColNm(whereCD8,i,rsSUM("col_"&i-1),rsSUM("col_"&i),"","","")
							ElseIf i = 3 Then
								sCodeName = db_GetSumColNm(whereCD8,i,rsSUM("col_"&i-2),rsSUM("col_"&i-1),rsSUM("col_"&i),"","")
							ElseIf i = 4 Then
								sCodeName = db_GetSumColNm(whereCD8,i,rsSUM("col_"&i-3),rsSUM("col_"&i-2),rsSUM("col_"&i-1),rsSUM("col_"&i),"")
							Else
								sCodeName = db_GetSumColNm(whereCD8,i,rsSUM("col_"&i-4),rsSUM("col_"&i-3),rsSUM("col_"&i-2),rsSUM("col_"&i-1),rsSUM("col_"&i))
							End if
							firstLine = firstLine & "<td align='center' class='TDCont' colspan='1' WIDTH='200'>"&sCodeName&"</td>"
						End If
					Next
	
					For i = sDepth + 1 To rsSUM.Fields.count
	
						firstLine = firstLine & "<td align='center' class='TDCont' colspan='1' WIDTH='200'>"&rsSUM("col_"&i)&"</td>"
	
					Next
					firstLine = "<tr bgcolor='"&sBG&"'>" & firstLine
					response.write firstLine & "</tr>"
	
					rsSUM.movenext
					'�Ұ�
				Loop
				%>
	
				</table>
				</div>
				
				<%
			end if
			%>
			</td>
		</tr>
	</table>
			
<% End if %>
<!-- #include virtual="/Include/Bottom.asp" -->