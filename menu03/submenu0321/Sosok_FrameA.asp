<!-- #include virtual="/Include/Top_Frame.asp" -->

<script>
<!--

// iframe 사이즈 적용

function fn_putetc2()
{
	try{
		eval("parent.document.all.whereCD5_A").value = document.all.whereCD5_A.value;
		eval("parent.document.all.whereCD5_B").value = "";
		eval("parent.document.all.whereCD5_C").value = "";
		eval("parent.document.all.whereCD5_D").value = "";
		eval("parent.document.all.whereCD5_E").value = "";
		parent.SosokFrameB.location = "/menu03/submenu0301/Sosok_FrameB.asp?SOSOK_A="+document.all.whereCD5_A.value+"&SOSOK_B=";
	}
	catch(e){}
}
-->
</script>

<!-- 프레임1 시작 -->


<%
SOSOK_A = Request("SOSOK_A")
%>
<body leftmargin="0" marginwidth="0" topmargin="0" marginheight="0" onLoad="ifHeight2('SosokFrameA');">

<div name="ifr" id="ifr">
<table cellspacing="0" cellpadding="0"  >
	<tr>
		<td>

<%
							'======= 처리구분 코드 가져오기 ==================================================
							SqlCode = "SELECT ACLASS CODE, CLASSNAME CODENAME FROM TB_ARMYINFO"
							SqlCode = SqlCode& " WHERE ACLASS < 'O' AND BCLASS IS NULL"
							SqlCode = SqlCode& " ORDER BY ACLASS"
							set RsCode = db.execute(SqlCode)
						%>
						&nbsp;<select name="whereCD5_A" size="1" class="ComboFFFCE7" onclick="fn_putetc2();">
							<Option value ='' selected>소속1차=====</option>
							<%
								IF NOT(RsCode.Eof OR RsCode.bof) THEN
									DO until RsCode.EOF
										CODE = RsCode("CODE")
										CODENAME = RsCode("CODENAME")
							%>
							<%=printSelect("" &CODENAME& "","" &CODE& "","" &SOSOK_A& "")%>
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
