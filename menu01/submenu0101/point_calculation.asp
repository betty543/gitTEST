<!-- #include virtual="/Include/Top2.asp" -->
<%
	On Error Resume next

	sControlName = Request("ControlName")
	iTot = Request("Tot")&".0"
	iCnt = Request("Cnt")
	sTot = Request("Value")


	if iTot <> "" then
		SQL = "select round("&iTot&"/"&iCnt&",2)"
		SET Rs = db.execute(SQL)

		sReturnValue = Rs(0)
		Rs.close
		SET Rs = nothing
	else

		sString = split(sTot,".")
		if mid(sString(1),3,1) <> "" and mid(sString(1),3,1) >= "5" then
			sReturnValue = sString(0)&"."&left(sString(1),2)+1
		else
			sReturnValue = sString(0)&"."&left(sString(1),2)
		end if

	end if


	response.write sString(0) &"<br>"
	sReturnValue = formatNumber(sReturnValue,2)

%>
	
	<script>
		eval("<%=sControlName%>").value = "<%=sReturnValue%>";
	</script>
<!-- #include virtual="/Include/Bottom.asp" -->