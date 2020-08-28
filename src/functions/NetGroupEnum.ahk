; ===============================================================================================================================
; Function ...: NetGroupEnum
; Return .....: Retrieves information about each global group in the security database, which is the security accounts manager (SAM)
;               database or, in the case of domain controllers, the Active Directory.
; Link .......: https://docs.microsoft.com/en-us/windows/win32/api/lmaccess/nf-lmaccess-netgroupenum
; ===============================================================================================================================

NetGroupEnum(ServerName := "127.0.0.1")
{
	static NERR_SUCCESS := 0
	static GROUP_INFO_1 := 1
	static MAX_PREFERRED_LENGTH := -1

	NET_API_STATUS := DllCall("netapi32\NetGroupEnum", "wstr",  ServerName
	                                                 , "uint",  GROUP_INFO_1
	                                                 , "ptr*",  buf
	                                                 , "uint",  MAX_PREFERRED_LENGTH
	                                                 , "uint*", EntriesRead
	                                                 , "uint*", TotalEntries
	                                                 , "uint*", 0
	                                                 , "uint")

	if (NET_API_STATUS = NERR_SUCCESS)
	{
		addr := buf, GROUP_INFO := []
		loop % EntriesRead
		{
			GROUP_INFO[A_Index, "name"]    := StrGet(NumGet(addr + A_PtrSize * 0, "uptr"), "utf-16")
			GROUP_INFO[A_Index, "comment"] := StrGet(NumGet(addr + A_PtrSize * 1, "uptr"), "utf-16")
			addr += A_PtrSize * 2
		}

		DllCall("netapi32\NetApiBufferFree", "ptr", buf, "uint")
		return GROUP_INFO
	}

	DllCall("netapi32\NetApiBufferFree", "ptr", buf, "uint")
	return false
}

; ===============================================================================================================================

NetGroupEnum := NetGroupEnum("DC01")
for i, v in NetGroupEnum {
	for k, v in NetGroupEnum[i]
		output .= k ": " v "`n"
	MsgBox % output
	output := ""
}