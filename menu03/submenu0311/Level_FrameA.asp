<!-- #include virtual="/Include/Top_Frame.asp" -->

<script>
<!--

// iframe ������ ����

function fn_putetc2()
{
	try{
		eval("parent.document.all.whereCD6_A").value = document.all.whereCD6_A.value;
		eval("parent.document.all.whereCD6_B").value = "";
		eval("parent.document.all.whereCD6_C").value = "";

		parent.LevelFrameB.location = "/menu03/submenu0301/Level_FrameB.asp?LEVEL_A="+document.all.whereCD6_A.value+"&LEVEL_B=";
	}
	catch(e){}
}
-->
</script>

<!-- ������1 ���� -->


<%
LEVEL_A = Request("LEVEL_A")
%>
<body leftmargin="0" marginwidth="0" topmargin="0" marginheight="0" onLoad="ifHeight2('LevelFrameA');">

<div name="ifr" id="ifr">
<table cellspacing="0" cellpadding="0"  >
	<tr>
		<td>

<%
							'======= ó������ �ڵ� �������� ==================================================
							SqlCode = "SELECT BCLASS CODE, CLASSNAME CODENAME FROM TB_ARMYINFO"
							SqlCode = SqlCode& " WHERE ACLASS = 'P' AND BCLASS IS NOT NULL AND CCLASS IS NULL"
							SqlCode = SqlCode& " ORDER BY BCLASS"
							set RsCode = db.execute(SqlCode)
						%>
						&nbsp;<select name="whereCD6_A" size="1" class="ComboFFFCE7" onclick="fn_putetc2();">
							<Option value ='' selected>���1��==</option>
							<%
								IF NOT(RsCode.Eof OR RsCode.bof) THEN
									DO until RsCode.EOF
										CODE = RsCode("CODE")
										CODENAME = RsCode("CODENAME")
							%>
							<%=printSelect("" &CODENAME& "","" &CODE& "","" &LEVEL_A& "")%>
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