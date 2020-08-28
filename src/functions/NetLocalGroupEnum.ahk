; ===============================================================================================================================
; Function ...: NetLocalGroupEnum
; Return .....: Information about each local group account on the specified server.
; Link .......: https://docs.microsoft.com/en-us/windows/win32/api/lmaccess/nf-lmaccess-netlocalgroupenum
; ===============================================================================================================================

NetLocalGroupEnum(ServerName := "127.0.0.1")
{
	static NERR_SUCCESS := 0
	static LOCALGROUP_INFO_1 := 1
	static MAX_PREFERRED_LENGTH := -1

	NET_API_STATUS := DllCall("netapi32\NetLocalGroupEnum", "wstr",  ServerName
	                                                      , "uint",  LOCALGROUP_INFO_1
	                                                      , "ptr*",  buf
	                                                      , "uint",  MAX_PREFERRED_LENGTH
	                                                      , "uint*", EntriesRead
	                                                      , "uint*", TotalEntries
	                                                      , "uint*", 0
	                                                      , "uint")

	if (NET_API_STATUS = NERR_SUCCESS)
	{
		addr := buf, LOCALGROUP_INFO := []
		loop % EntriesRead
		{
			LOCALGROUP_INFO[A_Index, "name"]    := StrGet(NumGet(addr + A_PtrSize * 0, "uptr"), "utf-16")
			LOCALGROUP_INFO[A_Index, "comment"] := StrGet(NumGet(addr + A_PtrSize * 1, "uptr"), "utf-16")
			addr += A_PtrSize * 2
		}

		DllCall("netapi32\NetApiBufferFree", "ptr", buf, "uint")
		return LOCALGROUP_INFO
	}

	DllCall("netapi32\NetApiBufferFree", "ptr", buf, "uint")
	return false
}

; ===============================================================================================================================

NetLocalGroupEnum := NetLocalGroupEnum("DC01")
for i, v in NetLocalGroupEnum {
	for k, v in NetLocalGroupEnum[i]
		output .= k ": " v "`n"
	MsgBox % output
	output := ""
}