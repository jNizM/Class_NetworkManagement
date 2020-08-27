; ===============================================================================================================================
; Function ...: NetWkstaUserEnum
; Return .....: Lists information about all users currently logged on to the workstation. This list includes interactive,
;               service and batch logons.
; Link .......: https://docs.microsoft.com/en-us/windows/win32/api/lmwksta/nf-lmwksta-netwkstauserenum
; ===============================================================================================================================

NetWkstaUserEnum(ServerName := "127.0.0.1")
{
	static NERR_SUCCESS := 0
	static ERROR_MORE_DATA := 234
	static WKSTA_USER_INFO_1 := 1
	static MAX_PREFERRED_LENGTH := -1

	NET_API_STATUS := DllCall("netapi32\NetWkstaUserEnum", "wstr",  ServerName
	                                                     , "uint",  WKSTA_USER_INFO_1
	                                                     , "ptr*",  buf
	                                                     , "uint",  MAX_PREFERRED_LENGTH
	                                                     , "uint*", EntriesRead
	                                                     , "uint*", TotalEntries
	                                                     , "uint*", 0
	                                                     , "uint")

	if ((NET_API_STATUS = NERR_SUCCESS) || (NET_API_STATUS = ERROR_MORE_DATA))
	{
		addr := buf, WKSTA_USER_INFO := []
		loop % EntriesRead
		{
			WKSTA_USER_INFO[A_Index, "username"]     := StrGet(NumGet(addr + A_PtrSize * 0, "uptr"), "utf-16")
			WKSTA_USER_INFO[A_Index, "logon_domain"] := StrGet(NumGet(addr + A_PtrSize * 1, "uptr"), "utf-16")
			WKSTA_USER_INFO[A_Index, "oth_domains"]  := StrGet(NumGet(addr + A_PtrSize * 2, "uptr"), "utf-16")
			WKSTA_USER_INFO[A_Index, "logon_server"] := StrGet(NumGet(addr + A_PtrSize * 3, "uptr"), "utf-16")
			addr += A_PtrSize * 4
		}

		DllCall("netapi32\NetApiBufferFree", "ptr", buf, "uint")
		return WKSTA_USER_INFO
	}

	DllCall("netapi32\NetApiBufferFree", "ptr", buf, "uint")
	return false
}

; ===============================================================================================================================

NetWkstaUserEnum := NetWkstaUserEnum()
for i, v in NetWkstaUserEnum {
	for k, v in NetWkstaUserEnum[i]
		output .= k ": " v "`n"
	MsgBox % output
	output := ""
}