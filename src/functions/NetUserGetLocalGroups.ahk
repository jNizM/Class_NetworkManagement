; ===============================================================================================================================
; Function ...: NetUserGetLocalGroups
; Return .....: Retrieves a list of local groups to which a specified user belongs.
; Link .......: https://docs.microsoft.com/en-us/windows/win32/api/lmaccess/nf-lmaccess-netusergetlocalgroups
; ===============================================================================================================================

NetUserGetLocalGroups(UserName, Domain := "", ServerName := "127.0.0.1")
{
	static NERR_SUCCESS := 0
	static LOCALGROUP_USERS_INFO_0 := 0
	static LG_INCLUDE_INDIRECT := 0x0001
	static MAX_PREFERRED_LENGTH := -1

	NET_API_STATUS := DllCall("netapi32\NetUserGetLocalGroups", "wstr",  ServerName
	                                                          , "wstr",  Domain . UserName
	                                                          , "uint",  LOCALGROUP_USERS_INFO_0
	                                                          , "uint",  LG_INCLUDE_INDIRECT
	                                                          , "ptr*",  buf
	                                                          , "uint",  MAX_PREFERRED_LENGTH
	                                                          , "uint*", EntriesRead
	                                                          , "uint*"  TotalEntries
	                                                          , "uint")

	if (NET_API_STATUS = NERR_SUCCESS)
	{
		addr := buf, LOCALGROUP_USERS_INFO := []
		loop % EntriesRead
		{
			LOCALGROUP_USERS_INFO.Push(StrGet(NumGet(addr + A_PtrSize * 0, "uptr"), "utf-16"))
			addr += A_PtrSize
		}

		DllCall("netapi32\NetApiBufferFree", "ptr", buf, "uint")
		return LOCALGROUP_USERS_INFO
	}

	DllCall("netapi32\NetApiBufferFree", "ptr", buf, "uint")
	return false
}

; ===============================================================================================================================

for k, v in NetUserGetLocalGroups(A_UserName, "DOMAIN\", "DC01")
	output .= k ": " v "`n"
MsgBox % output