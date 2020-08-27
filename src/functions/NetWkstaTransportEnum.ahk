; ===============================================================================================================================
; Function ...: NetWkstaTransportEnum
; Return .....: Supplies information about transport protocols that are managed by the redirector, which is the software on the
;               client computer that generates file requests to the server computer.
; Link .......: https://docs.microsoft.com/en-us/windows/win32/api/lmwksta/nf-lmwksta-netwkstatransportenum
; ===============================================================================================================================

NetWkstaTransportEnum(ServerName := "127.0.0.1")
{
	static NERR_SUCCESS := 0
	static ERROR_MORE_DATA := 234
	static WKSTA_TRANSPORT_INFO_0 := 0
	static MAX_PREFERRED_LENGTH := -1

	NET_API_STATUS := DllCall("netapi32\NetWkstaTransportEnum", "wstr",  ServerName
	                                                          , "uint",  WKSTA_TRANSPORT_INFO_0
	                                                          , "ptr*",  buf
	                                                          , "uint",  MAX_PREFERRED_LENGTH
	                                                          , "uint*", EntriesRead
	                                                          , "uint*", TotalEntries
	                                                          , "uint*", 0
	                                                          , "uint")

	if ((NET_API_STATUS = NERR_SUCCESS) || (NET_API_STATUS = ERROR_MORE_DATA))
	{
		addr := buf, WKSTA_TRANSPORT_INFO := []
		loop % EntriesRead
		{
			WKSTA_TRANSPORT_INFO[A_Index, "quality_of_service"] := NumGet(addr + 0, "uint")
			WKSTA_TRANSPORT_INFO[A_Index, "number_of_vcs"]      := NumGet(addr + 4, "uint")
			WKSTA_TRANSPORT_INFO[A_Index, "transport_name"]     := StrGet(NumGet(addr + A_PtrSize * 2, "uptr"), "utf-16")
			WKSTA_TRANSPORT_INFO[A_Index, "transport_address"]  := StrGet(NumGet(addr + A_PtrSize * 3, "uptr"), "utf-16")
			WKSTA_TRANSPORT_INFO[A_Index, "wan_ish"]            := NumGet(addr + A_PtrSize * 4, "int")
			addr += 8 + A_PtrSize * 3
		}

		DllCall("netapi32\NetApiBufferFree", "ptr", buf, "uint")
		return WKSTA_TRANSPORT_INFO
	}

	DllCall("netapi32\NetApiBufferFree", "ptr", buf, "uint")
	return false
}

; ===============================================================================================================================

NetWkstaTransportEnum := NetWkstaTransportEnum()
for i, v in NetWkstaTransportEnum {
	for k, v in NetWkstaTransportEnum[i]
		output .= k ": " v "`n"
	MsgBox % output
	output := ""
}