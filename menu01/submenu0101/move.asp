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
			'��� ->��������
			SQL = "UPDATE armyinformix.dbo.receiptfact SET  processgb = '1', processdate= null, monitorpoint = null WHERE receiptfactnum = '" & factnum & "'"
			db.execute(SQL)
		elseif FRM = "submenu02" then
			'��� ->��������
			SQL = "UPDATE armyinformix.dbo.receiptfact SET  processgb = null, processdate= null, monitorpoint = null WHERE receiptfactnum = '" & factnum & "'"
			db.execute(SQL)
		elseif FRM = "submenu05" then
			'�Ϸ���� ����ڷῡ�� ������Ű��
			SQL = "UPDATE armyinformix.dbo.receiptfact SET  processgb = '8' WHERE receiptfactnum = '" & factnum & "'"
			db.execute(SQL)
		elseif FRM = "submenu03" then
			'����ڷ� ������ ����ڷ�� ���Խ�Ű��
			SQL = "UPDATE armyinformix.dbo.receiptfact SET  processgb = '9' WHERE receiptfactnum = '" & factnum & "'"
			db.execute(SQL)
		end if

		response.write SQL

	NEXT
END IF



%>
	
<!-- #include virtual="/Include/Bottom.asp" -->