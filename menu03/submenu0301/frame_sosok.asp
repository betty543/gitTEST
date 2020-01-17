<!-- #include virtual="/Include/Top_Frame.asp" -->

<script>
<!--

// iframe 사이즈 적용
function fn_SetSosok3(arg,arg1)
{
	parent.frame_sosok2.location = "/menu03/submenu0301/frame_sosok_3.asp?SOSOKGB="+arg+"&SOSOKETCGB="+arg1;

}
function fn_putetc2()
{
	try{
		//eval("parent.document.all.whereCD7").value = document.all.level2.value;
		eval("parent.document.all.SOSOKETCGB").value = document.all.level2.value;
		parent.frame_sosok2.location = "/menu03/submenu0301/frame_sosok_3.asp?SOSOKGB="+parent.document.all.SOSOKGB.value+"&SOSOKETCGB="+document.all.level2.value;
	}
	catch(e){}
}
-->
</script>

<!-- 프레임1 시작 -->
<form name="frmCode" style="margin:0">

<%
SOSOKGB = Request("SOSOKGB")
SOSOKETCGB = Request("SOSOKETCGB")
%>
<body bgcolor="#FFFFFF" topmargin=0 leftmargin=0>
<table width="100%" height="100%" border="0" cellspacing="0" cellpadding="0" align=left  bgcolor="#000000">
   <tr>
    <td bgcolor="FFFFFF" align="left" height="29" valign="center">

		<%
			if SOSOKGB = "" then
				sReplyHtml = "소속(대)를 선택하시면 중분류가 표시됩니다."
				response.write sReplyHtml
			else
				'======= 처리구분 코드 가져오기 ==================================================
				SqlCode = "select * from tb_armyinfo where aclass = '"&SOSOKGB&"' and bclass is not null and cclass is null order by bclass"
				set RsCode = db.execute(SqlCode)

				do until RsCode.eof
					j = j + 1
					SelectedValue = ""
					if j = 1 then
						sReplyHtml = "<input type='RADIO' value='" & RsCode("bclass") & "' name='SOSOKGB2' class='none' " & SelectedValue & " onClick=""fn_SetSosok3('"&RsCode("aclass")&"','"&RsCode("bclass")&"');"">" & RsCode("classname")	
					else
						sReplyHtml = sReplyHtml & "&nbsp;<input type='RADIO' value='" & RsCode("bclass") & "' name='SOSOKGB2' class='none' " & SelectedValue & " onClick=""fn_SetSosok3('"&RsCode("aclass")&"','"&RsCode("bclass")&"');"">" & RsCode("classname")	
					end if
					RsCode.movenext
				loop
				RsCode.close
				response.write sReplyHtml

			end if
		%>						

	</td>
</tr>
</table>
</form>
<!-- #include virtual="/Include/Bottom.asp" -->
