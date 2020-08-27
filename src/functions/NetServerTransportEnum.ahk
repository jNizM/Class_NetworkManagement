; ===============================================================================================================================
; Function ...: NetServerTransportEnum
; Return .....: Supplies information about transport protocols that are managed by the server.
; Link .......: https://docs.microsoft.com/en-us/windows/win32/api/lmserver/nf-lmserver-netservertransportenum
; ===============================================================================================================================

NetServerTransportEnum(ServerName := "127.0.0.1")
{
	static NERR_SUCCESS := 0
	static ERROR_MORE_DATA := 234
	static SERVER_TRANSPORT_INFO_1 := 1
	static MAX_PREFERRED_LENGTH := -1

	NET_API_STATUS := DllCall("netapi32\NetServerTransportEnum", "wstr",  ServerName
	                                                           , "uint",  SERVER_TRANSPORT_INFO_1
	                                                           , "ptr*",  buf
	                                                           , "uint",  MAX_PREFERRED_LENGTH
	                                                           , "uint*", EntriesRead
	                                                           , "uint*", TotalEntries
	                                                           , "uint*", 0
	                                                           , "uint")

	if ((NET_API_STATUS = NERR_SUCCESS) || (NET_API_STATUS = ERROR_MORE_DATA))
	{
		addr := buf, SERVER_TRANSPORT_INFO := []
		loop % EntriesRead
		{
			SERVER_TRANSPORT_INFO[A_Index, "numberofvcs"]            := NumGet(addr + 0, "uint")
			SERVER_TRANSPORT_INFO[A_Index, "transportname"]          := StrGet(NumGet(addr + A_PtrSize * 1, "uptr"), "utf-16")
			SERVER_TRANSPORT_INFO[A_Index, "transportaddress"]       := NumGet(addr + A_PtrSize * 2, "ptr")   ; todo
			SERVER_TRANSPORT_INFO[A_Index, "transportaddresslength"] := NumGet(addr + A_PtrSize * 3, "uint")
			SERVER_TRANSPORT_INFO[A_Index, "networkaddress"]         := StrGet(NumGet(addr + A_PtrSize * 4, "uptr"), "utf-16")
			SERVER_TRANSPORT_INFO[A_Index, "domain"]                 := StrGet(NumGet(addr + A_PtrSize * 5, "uptr"), "utf-16")
			addr += A_PtrSize * 6
		}

		DllCall("netapi32\NetApiBufferFree", "ptr", buf, "uint")
		return SERVER_TRANSPORT_INFO
	}

	DllCall("netapi32\NetApiBufferFree", "ptr", buf, "uint")
	return false
}

; ===============================================================================================================================

NetServerTransportEnum := NetServerTransportEnum()
for i, v in NetServerTransportEnum {
	for k, v in NetServerTransportEnum[i]
		output .= k ": " v "`n"
	MsgBox % output
	output := ""
}