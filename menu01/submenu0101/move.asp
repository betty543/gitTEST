<!-- #include virtual="/Include/Top2.asp" -->

<%
	On Error Resume next


	FRM = Request("FRM")
	factnumlist = Request("factnum")


If factnumlist <> "" and instr(factnumlist,",")>0 then

	selectvalue_value = Split(factnumlist,",")
	i  = 0
	FOR k = 1 to UBound(selectvalue_value)


		factnum = selectvalue_value(k)

		if FRM = "submenu01" then
			'대기 ->진행으로
			SQL = "UPDATE armyinformix.dbo.receiptfact SET  processgb = '1', processdate= null, monitorpoint = null WHERE receiptfactnum = '" & factnum & "'"
			db.execute(SQL)
		elseif FRM = "submenu02" then
			'대기 ->진행으로
			SQL = "UPDATE armyinformix.dbo.receiptfact SET  processgb = null, processdate= null, monitorpoint = null WHERE receiptfactnum = '" & factnum & "'"
			db.execute(SQL)
		elseif FRM = "submenu05" then
			'완료건을 통계자료에서 누락시키기
			SQL = "UPDATE armyinformix.dbo.receiptfact SET  processgb = '8' WHERE receiptfactnum = '" & factnum & "'"
			db.execute(SQL)
		elseif FRM = "submenu03" then
			'통계자료 누락건 통계자료로 포함시키기
			SQL = "UPDATE armyinformix.dbo.receiptfact SET  processgb = '9' WHERE receiptfactnum = '" & factnum & "'"
			db.execute(SQL)
		end if

		response.write SQL

	NEXT
END IF



%>
	
<!-- #include virtual="/Include/Bottom.asp" -->