<!--#INCLUDE FILE="database.asp" -->
<%
    Dim bDisableUTF8, myDoNotCheckAccess, sApplicationRoot, sVirtualRoot, strUserAgent

    Rem -- Force UTF-8
    If Not IsEmpty(Session) Then
        If Not CBoolEx(bDisableUTF8) Then
            Session.CodePage = 65001 ' UTF-8
        Else        
            Session.CodePage = 1252 ' Windows-1252
        End If
    Else
        If Not CBoolEx(bDisableUTF8) Then
            Response.Codepage = 65001 
            Response.Charset = "utf-8" 
        End If
    End If

	Const VersionNo = "3.03"
	Const MajorVersion = 3
	Const ServicePackVersion = "00"
	Dim MainTitle, VersionTxt
	MainTitle = "ExpandIT Mobile " & VersionNo
	VersionTxt = "3.03"
	
	Rem - Inserted to prevent error when using Response.Flush (2002-06-21)
	Rem ------------------------------------------------------------------
    On Error Resume Next
	Response.Buffer = True
	Err.Clear
	On Error Goto 0
	Rem ------------------------------------------------------------------
	
		
	Dim conn
	dim dbDatabase
	
	set dbDatabase = nothing


	Rem -- Prior to 2.12, connection caching was implemented
	Rem -- This has been disabled in 2.13
	Set conn = CreateObject("ADODB.Connection")
	conn.CommandTimeout = 900
	conn.ConnectionTimeout = 30
	conn.Open ConnectionString

    Rem -- Load the session from the database if page is running with EnableSessionState=False
    If IsEmpty(Session) Then Set Session = LoadSession()


	REM - ********************************************************************************************************
	REM - * This will load the user data, which will be used when checking for website access.
	REM - * 
	REM - * If you have a page on which you do not want a check. Then apply the following line before including
	REM - * the files "common.asp" or "header.asp".
	REM - *
	REM - * myDoNotCheckAccess = True
	REM - *
	REM - ********************************************************************************************************
    If Not myDoNotCheckAccess Then
		LoadAdminUser
	End If
	REM - ********************************************************************************************************

    Function HandleCounters(dicCounters, rsClient, dicRetCounters)
        Dim i, CounterName, CounterVal
        
        For i = 0 To rsClient.Fields.Count - 1
            CounterName = rsClient.Fields(i).Name
            Select Case CounterName
                Case "BASClient"
                Case "BASProfileGuid"
                Case "Description"
                Case "Owner"
                Case "UserGuid"
                Case "NTLMName"
                Case "UserLogin"
                Case "Password"
                Case "Blocked"
                Case "DeviceClass"
                Case "DeviceName"
                Case "OSName"
                Case "OSVersion"
                Case "ClientType"
                Case "ClientProcessor":
                        Rem -- Do not treat as counter
                Case Else
                        Rem -- This is a counter
                        Rem -- Removed decimals and dates: 5, 6, 7, 14, 131, 133, 134, 135
                        Select Case rsClient.Fields(i).Type
                            Case 2, 3, 4, 16, 17, 18, 19, 20, 21
                                Rem -- this is a numeric value and should be treated as a counter
                                If CLngEx(dicCounters(CounterName)) <= 0 Or Abs(CLngEx(dicCounters(CounterName))) > 1000000 Then
                                    Rem -- An invalid counter value was received...
                                    CounterVal = CLngEx(rsClient(CounterName).Value)
                                Else
                                    Rem -- Select the highest counter (Max wins)
                                    If CLngEx(dicCounters(CounterName)) > CLngEx(rsClient(CounterName).Value) Then
                                        CounterVal = CLngEx(dicCounters(CounterName))
                                    Else
                                        CounterVal = CLngEx(rsClient(CounterName).Value)
                                    End If
                                End If
                                Rem -- Finally, write the changes back
                                rsClient(CounterName).Value = CounterVal
                                If TypeName(dicRetCounters) = "IDictionary" Then
                                    dicRetCounters(CounterName) = CounterVal
                                End If
                                
                            Case Else
                                Rem -- Ignore it..
                        End Select
            End Select
        Next
        rsClient.Update
    End Function


	Rem -- Backwards compatibility
	Function HandleCounter(CounterName, dicCounters, rsClient)
		dim retval
		
		If TypeName(dicCounters) <> "IDictionary" Then
		    If IsNull(rsClient("OrderCounter").Value) Then
		        retval = 1
		    Else
		        retval = CLng(rsClient(CounterName).Value)
		    End If
		Else
			'If doDebug Then 
			'	If CLng(abs(dicCounters.OrderCounter)) > 1000000 Then
			'		Response.Write "Kammerater! Så er den gal igen! Tallet er [" & CLng(abs(receiveddict.Counters.OrderCounter)) & "]"
			'		Response.End
			'	End If
			'End If
			
			'New Clients creates a really big ordercounter - if counter is larger than 1000000 then get the number from SQL Server
		    If CLng(dicCounters(CounterName)) <= 0 Or CLng(abs(dicCounters(CounterName))) > 1000000  Then
		        retval = CLng(rsClient("OrderCounter").Value)
		    Else
		        retval = CLng(dicCounters(CounterName))
		    End If
		End If
		HandleCounter = retval
	End Function
	
	
	sApplicationRoot = GetShopRoot()
	sVirtualRoot = GetVirtualRoot(sApplicationRoot)
	Application("VRoot") = sVirtualRoot
			

	Rem -- *************************************************************************
	Rem -- GetVirtualRoot finds the virtual root of the shop when given the 
	Rem -- physical path to the root of the shop.
	Rem -- *************************************************************************
	Function GetVirtualRoot(approot)
		Dim f
		Dim vroot, p, c
		vroot = Request.ServerVariables("PATH_INFO")
		approot = ucase(approot)
		Do
			If Ucase(Server.MapPath(vroot)) = approot Then Exit Do
			While (Right(vroot,1) <> "/") And (Len(vroot)>1)
				vroot = Mid(vroot,1, Len(vroot)-1)
			Wend
			vroot = Mid(vroot,1, Len(vroot)-1)
		Loop While Len(vroot)>1
		If vroot = "/" Then vroot = ""
		GetVirtualRoot = vroot
	End Function


	Function GetShopRoot()
		Dim p, c, root
		root = Request.ServerVariables("APPL_PHYSICAL_PATH") 'Server.MapPath("/")
		'root = Server.MapPath("/")
		p = Server.MapPath(".")
		c = 20
		Do	
			c = c - 1
			If c=0 Then Exit Do
				
			If Not IsRoot(p) Then
				While Len(p) > Len(root) And Mid(p, Len(p), 1) <> "\"
					p = Mid(p,1,Len(p)-1)
				Wend
				If Mid(p, Len(p), 1) = "\" Then
					p = Mid(p,1,Len(p)-1)
				End If			
			Else 
				Exit Do
			End If
		Loop 
			
		GetShopRoot = p
	End Function
	
	Rem -- *************************************************************************
	Rem -- IsRoot determines If the path specified is the root path of the shop 
	Rem -- installation.
	Rem -- 
	Rem -- It doese so by attempting to open the file "global.mdb" as a database. 
	Rem -- This database must be in the root of the shop catalog. 
	Rem -- *************************************************************************
	Function IsRoot(path)
		Dim conn, retv, errdesc, errnum
		
		retv = False
		Set conn = Server.CreateObject("ADODB.Connection")
		On Error Resume Next
		conn.Open 	"DRIVER={Microsoft Access Driver (*.mdb)};" & _
					"DBQ=" & path & "\global.mdb;" & _
					"SAVEFILE=;"
		errnum = Err.Number
		errdesc = Ucase(Err.Description)
		On Error Goto 0
		If errnum <> 0 Then
			retv = false 
		ElseIf errnum = 0 And conn.state = 1 Then
			retv = true
		Else
			Response.Write GetLabel(MESSAGE_UNABLE_TO_DETECT_ROOT_DIRECTORY) & ".<P>"
			Response.End
		End If
		
		IsRoot = retv
	End Function
%>
<!-- #INCLUDE FILE="include/util.asp" -->
