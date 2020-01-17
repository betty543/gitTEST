<!-- #include virtual="/Include/Top_Frame.asp" -->

<script>
<!--

// iframe 사이즈 적용

function fn_putetc2()
{
	try{
		eval("parent.document.all.whereCD13_A").value = document.all.whereCD13_A.value;
		eval("parent.document.all.whereCD13_B").value = "";

		parent.CallClassFrameB.location = "/menu03/submenu0321/CallClass_FrameB.asp?CALLCLASS_A="+document.all.whereCD13_A.value+"&CALLCLASS_B=";
	}
	catch(e){}
}
-->
</script>

<!-- 프레임1 시작 -->


<%
CALLCLASS_A = Request("CALLCLASS_A")
%>
<body leftmargin="0" marginwidth="0" topmargin="0" marginheight="0" onLoad="ifHeight2('CallClassFrameA');">

<div name="ifr" id="ifr">
<table cellspacing="0" cellpadding="0"  >
	<tr>
		<td>

<%
							'======= 처리구분 코드 가져오기 ==================================================
							SqlCode = "SELECT BCLASS CODE, CLASSNAME CODENAME FROM TB_ARMYINFO"
							SqlCode = SqlCode& " WHERE ACLASS = 'S' AND BCLASS IS NOT NULL AND CCLASS IS NULL"
							SqlCode = SqlCode& " ORDER BY BCLASS"
							set RsCode = db.execute(SqlCode)
						%>
						&nbsp;<select name="whereCD13_A" size="1" class="ComboFFFCE7" onclick="fn_putetc2();">
							<Option value ='' selected>상담분야1차==</option>
							<%
								IF NOT(RsCode.Eof OR RsCode.bof) THEN
									DO until RsCode.EOF
										CODE = RsCode("CODE")
										CODENAME = RsCode("CODENAME")
							%>
							<%=printSelect("" &CODENAME& "","" &CODE& "","" &CALLCLASS_A& "")%>
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
