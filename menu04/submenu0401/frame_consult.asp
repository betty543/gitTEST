<!-- #include virtual="/Include/Top_Frame.asp" -->

<script>
<!--

// iframe ������ ����
function ifHeight2(f){
	obj = eval("parent.document.all."+ f);
	obj.style.height = document.all.level2.offsetHeight;
	obj.style.width = document.all.level2.offsetWidth;
}


function fn_putetc2()
{
	//try{
		//eval("parent.document.all.whereCD7").value = document.all.level2.value;
		eval("parent.document.inUpFrm.CONSULTETCGB").value = document.all.level2.value;
	//}
	//catch(e){}
}
-->
</script>

<!-- ������1 ���� -->
<form name="frmCode" style="margin:0">

<%
CONSULTGB = Request("CONSULTGB")
CONSULTETCGB = Request("CONSULTETCGB")
%>
<body bgcolor="#FFFFFF" topmargin=0 leftmargin=0 onload="ifHeight2('frame_consult');">
<table width="100" height="100%" border="0" cellspacing="0" cellpadding="0" align=left  bgcolor="#000000">
   <tr>
    <td bgcolor="FFFFFF" align="left" height="29" valign="center"><select name="level2" onchange="javascript:fn_putetc2();">		
		<option value="" <% if CONSULTETCGB = "" then%>selected<%end if%>>��Ÿ�о�</option>

		<%
			if CONSULTGB = "Z" then
				'======= ó������ �ڵ� �������� ==================================================
				SqlCode = "SELECT CODE, CODENAME FROM TB_CODE"
				SqlCode = SqlCode& " WHERE USEYN='Y' AND SYSYN='Y' AND CODEGROUP='C31'"
				SqlCode = SqlCode& " ORDER BY CODE"
				set RsCode = db.execute(SqlCode)

				do until RsCode.eof

					CODE = RsCode("CODE")
					CODENAME = RsCode("CODENAME")
							%>
							<%=printSelect("" &CODENAME& "","" &CODE& "","" &CONSULTETCGB& "")%>
							<%
					RsCode.MoveNext
				LOOP

				RsCode.Close
				set RsCode = NOTHING
			end if
		%></select>
	</td>
</tr>
</table>
</form>
<!-- #include virtual="/Include/Bottom.asp" -->