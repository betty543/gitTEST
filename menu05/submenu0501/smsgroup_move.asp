<!-- #include virtual="/Include/Top_Frame.asp" -->

<%
INCODE = SESSION("SS_LoginID")
selectvalue		= Request("selectvalue")
ogroupidx		= Request("ogroupidx")
ngroupidx		= Request("ngroupidx")
JOBGB = Request("JOBGB") 

	'	response.write ngroupidx

If selectvalue <> "" and instr(selectvalue,":")>0 then

	selectvalue_value = Split(selectvalue,":")
	i  = 0
	FOR k = 1 to UBound(selectvalue_value)

		strSQL2 = "SELECT name,cellphone,class,armyno,sosok_id,sosok_name FROM TB_SMSADDR where idx = "& selectvalue_value(k) &""
		'response.write strSQL2


		Set rss = DB.Execute(strsql2)
		if Not(rss.eof or rss.bof) Then

			Do While Not RSs.EOF
				sname	= rss("name")
				cellphone	= rss("cellphone")
				sclass	= rss("class")
				armyno	= rss("armyno")
				sosok_id	= rss("sosok_id")
				sosok_name	= rss("sosok_name")

				strSQL = "SELECT * FROM TB_SMSADDR WHERE group_idx = '"& ngroupidx & "' and cellphone = '" & cellphone & "'"

				'존재하면
				set rs2 = DB.Execute(strSQL)	
				if rs2.eof then

					strSQL = "INSERT INTO TB_SMSADDR ( group_idx, name,cellphone,class,armyno,sosok_id,sosok_name,incode,indate,mocode,modate)" &_
						" values ("&ngroupidx&",'"& sname	& "', " &_
								"'" & cellphone		& "', '" & sclass		& "', '" & armyno		& "', '" & sosok_id		& "', '" & sosok_name		& "', '" & INCODE		& "',getdate(), '" & INCODE		& "',getdate())"

					response.write strSQL
					DB.Execute(strSQL)

				else
					i  = i + 1
				end if

				if JOBGB = "MOV" then
					strSQL = "DELETE FROM TB_SMSADDR WHERE idx = "& selectvalue_value(k) &""
					DB.Execute(strSQL)
				end if
				rss.MoveNext
			Loop


		end if
	NEXT

End If 

if JOBGB = "MOV" then
%>	
		<SCRIPT LANGUAGE="JavaScript">
		<!--
			alert("!!! 이동을 완료 하였습니다. 총 [ <%=k-1%> ] 건 // 중복 [ <%=i%> ] 건이 제외되었습니다.!!!");
			self.location.href = 'smsgroup.asp';
		//-->
		</SCRIPT>

<% else %>

		<SCRIPT LANGUAGE="JavaScript">
		<!--
			alert("!!! 추가를 완료 하였습니다. 총 [ <%=k-1%> ] 건 // 중복 [ <%=i%> ] 건이 제외되었습니다.!!!");

			self.location.href = 'smsgroup.asp';
		//-->
		</SCRIPT>

<% end if %>