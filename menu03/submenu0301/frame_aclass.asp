<!-- #include virtual="/Include/Top_Frame.asp" -->

<script>
<!--

// iframe ������ ����
function fn_putetc3(arg)
{
	try{
		if ( eval("parent.document.all.whereCD7") != null )
			eval("parent.document.all.whereCD7").value = arg;

		eval("parent.document.all.LEVEL2").value = arg;
	}
	catch(e){}
}
-->
</script>

<!-- ������1 ���� -->
<form name="frmCode" style="margin:0">

<%
level = Request("level")
level2 = Request("level2")

%>
<body bgcolor="#FFFFFF" topmargin= leftmargin=0>
<table width="100%" height="100%" border="0" cellspacing="0" cellpadding="0" align=left  bgcolor="#000000">
   <tr>
    <td bgcolor="FFFFFF" align="left" height="29" valign="center">


		<%if level = "A" or  level = "B" then 

				'======= ó������ �ڵ� �������� ==================================================
				SqlCode = "SELECT CODE, CODENAME FROM TB_CODE"
				SqlCode = SqlCode& " WHERE USEYN='Y' AND SYSYN='Y' AND CODEGROUP='C06'"
				SqlCode = SqlCode& " ORDER BY CODE"
				set RsCode = db.execute(SqlCode)
				do until RsCode.eof
					j = j + 1
					SelectedValue = ""
					if j = 1 then
						sReplyHtml = "<input type='RADIO' value='" & RsCode("CODE") & "' name='BCLASS' class='none' " & SelectedValue & "  onClick=""fn_putetc3('"&RsCode("CODE")&"');"">" & RsCode("CODENAME")	
					else
						sReplyHtml = sReplyHtml & "&nbsp;<input type='RADIO' value='" & RsCode("CODE") & "' name='BCLASS' class='none' " & SelectedValue & "  onClick=""fn_putetc3('"&RsCode("CODE")&"');"">" & RsCode("CODENAME")	
					end if
					RsCode.movenext
				loop
				RsCode.close
				response.write sReplyHtml
		else

				sReplyHtml = "���о�(��)�� �����Ͻø� �ߺз��� ǥ�õ˴ϴ�."
				response.write sReplyHtml

		end if
		%>
	</td>
</tr>
</table>
</form>
<!-- #include virtual="/Include/Bottom.asp" -->
