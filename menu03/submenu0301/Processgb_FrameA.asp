<!-- #include virtual="/Include/Top_Frame.asp" -->

<script>
<!--

// iframe 사이즈 적용

function fn_putetc2()
{
	try{
		eval("parent.document.all.whereCD14").value = document.all.whereCD14.value;
		eval("parent.document.all.whereCD14_B").value = "";

		parent.ProcessgbFrameB.location = "/menu03/submenu0301/Processgb_FrameB.asp?PROCESSGB_A="+document.all.whereCD14.value+"&PROCESSGB_B=";
	}
	catch(e){}
}
-->
</script>

<!-- 프레임1 시작 -->


<%
PROCESSGB_A = Request("PROCESSGB_A")
%>
<body leftmargin="0" marginwidth="0" topmargin="0" marginheight="0" onLoad="ifHeight2('ProcessgbFrameA');">

<div name="ifr" id="ifr">
<table cellspacing="0" cellpadding="0"  >
	<tr>
		<td>

<%
							'======= 상담유형 코드 가져오기 ==================================================
						SqlCode = "SELECT BCLASS CODE, CLASSNAME CODENAME FROM TB_ARMYINFO"
						SqlCode = SqlCode& " WHERE ACLASS = 'U' AND BCLASS IS NOT NULL  AND CCLASS IS NULL"
						SqlCode = SqlCode& " ORDER BY BCLASS"
						set RsCode = db.execute(SqlCode)
							set RsCode = db.execute(SqlCode)
						%>
						&nbsp;<select name="whereCD14" size="1" class="ComboFFFCE7" onclick="fn_putetc2();">
							<Option value ='' selected>조치결과1차==</option>
							<%
								IF NOT(RsCode.Eof OR RsCode.bof) THEN
									DO until RsCode.EOF
										CODE = RsCode("CODE")
										CODENAME = RsCode("CODENAME")
							%>
							<%=printSelect("" &CODENAME& "","" &CODE& "","" &PROCESSGB_A& "")%>
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
