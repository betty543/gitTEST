<!-- #include virtual="/Include/Top_Frame.asp" -->

<%
INCODE = SESSION("SS_LoginID")
FRM		= Request("FRM")
idx		= Request("idx")

	'	response.write ngroupidx

if FRM = "right" then
	'3자통화목록에 추가


		strSQL2 = "SELECT name,gunphone,cellphone,class,armyno,sosok_id,sosok_name FROM TB_SMSADDR where idx = "& idx &""
		'response.write strSQL2


		Set rss = DB.Execute(strsql2)
		if Not(rss.eof or rss.bof) Then

			Do While Not RSs.EOF
				sname	= rss("name")
				cellphone	= rss("cellphone")
				gunphone	= rss("gunphone")
				sclass	= rss("class")
				armyno	= rss("armyno")
				sosok_id	= rss("sosok_id")
				sosok_name	= rss("sosok_name")

				strSQL = "SELECT * FROM temp_conference WHERE addr_idx = '"& idx & "' and datagb = '2'"
				'존재하면
				set rs2 = DB.Execute(strSQL)	
				if rs2.eof then

					strSQL = "INSERT INTO temp_conference ( addr_idx, userid, cellphone, gunphone, datagb)" &_
						" values ("&idx&",'"& INCODE	& "', " &_
								"'" & cellphone		& "','" & gunphone		& "','2')"

					'response.write strSQL
					DB.Execute(strSQL)

				else
					i  = i + 1
				end if


				rss.MoveNext
			Loop


		end if
	'NEXT
else
	strSQL = "DELETE FROM temp_conference WHERE idx = '"& idx & "'"
	DB.Execute(strSQL)

	RESPONSE.WRITE strSQL
End If 

%>