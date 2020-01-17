<!-- #include virtual="/Include/Top_Frame.asp" -->

<%
INCODE = SESSION("SS_LoginID")
FRM		= Request("FRM")
factnumlist		= Request("idx")

	'	response.write ngroupidx

if FRM = "right" then
	'3자통화목록에 추가


	If factnumlist <> "" and instr(factnumlist,",")>0 then

		selectvalue_value = Split(factnumlist,",")
		i  = 0
		FOR k = 1 to UBound(selectvalue_value)

			idx = selectvalue_value(k)


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

					strSQL = "SELECT * FROM temp_conference WHERE addr_idx = '"& idx & "' and datagb = '1'"
					'존재하면
					set rs2 = DB.Execute(strSQL)	
					if rs2.eof then

						strSQL = "INSERT INTO temp_conference ( addr_idx, userid, cellphone, gunphone, datagb, successflag)" &_
							" values ("&idx&",'"& INCODE	& "', " &_
									"'" & cellphone		& "','" & gunphone		& "','1','0')"

						'response.write strSQL
						DB.Execute(strSQL)

					else
						i  = i + 1
					end if


					rss.MoveNext
				Loop

			end if
		NEXT
	end if
else


	If factnumlist <> "" and instr(factnumlist,",")>0 then

		selectvalue_value = Split(factnumlist,",")
		i  = 0
		FOR k = 1 to UBound(selectvalue_value)

			idx = selectvalue_value(k)

			strSQL = "DELETE FROM temp_conference WHERE idx = '"& idx & "'"
			DB.Execute(strSQL)

			'RESPONSE.WRITE strSQL
		next
	end if
End If 

%>