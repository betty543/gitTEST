<!-- #include virtual="/Include/Top_Frame.asp" -->

<script>
<!--

// iframe ������ ����

function fn_putetc2()
{
	try{
		eval("parent.document.all.whereCD5_C").value = document.all.whereCD5_C.value;
		eval("parent.document.all.whereCD5_D").value = "";
		eval("parent.document.all.whereCD5_E").value = "";
		parent.SosokFrameD.location = "/menu03/submenu0301/Sosok_FrameD.asp?SOSOK_A="+eval("parent.document.all.whereCD5_A").value+"&SOSOK_B="+eval("parent.document.all.whereCD5_B").value+"&&SOSOK_C="+eval("parent.document.all.whereCD5_C").value+"&SOSOK_D=";
	}
	catch(e){}
}
-->
</script>

<!-- ������1 ���� -->


<%
SOSOK_A = Request("SOSOK_A")
SOSOK_B = Request("SOSOK_B")
SOSOK_C = Request("SOSOK_C")
%>
<body leftmargin="0" marginwidth="0" topmargin="0" marginheight="0" onLoad="ifHeight2('SosokFrameC');">

<div name="ifr" id="ifr">
<table cellspacing="0" cellpadding="0"  >
	<tr>
		<td>

<%
							'======= ó������ �ڵ� �������� ==================================================
							SqlCode = "SELECT CCLASS CODE, CLASSNAME CODENAME FROM TB_ARMYINFO"
							SqlCode = SqlCode& " WHERE ACLASS = '" &SOSOK_A&"' AND BCLASS = '" &SOSOK_B&"' AND CCLASS IS NOT NULL AND DCLASS IS NULL"
							SqlCode = SqlCode& " ORDER BY CCLASS"
							set RsCode = db.execute(SqlCode)
						%>
						&nbsp;<select name="whereCD5_C" size="1" class="ComboFFFCE7" onclick="fn_putetc2();">
							<Option value ='' selected>�Ҽ�3��=====</option>
							<%
								IF NOT(RsCode.Eof OR RsCode.bof) THEN
									DO until RsCode.EOF
										CODE = RsCode("CODE")
										CODENAME = RsCode("CODENAME")
							%>
							<%=printSelect("" &CODENAME& "","" &CODE& "","" &SOSOK_C& "")%>
							<%
									RsCode.MoveNext
									LOOP
								END IF
								RsCode.Close
								set RsCode = NOTHING
							%>
						</select>						

		</td>
	</tr>
</table>
</div>

<!-- #include virtual="/Include/Bottom.asp" -->
