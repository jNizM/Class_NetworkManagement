; ===============================================================================================================================
; Function ...: NetUserGetGroups
; Return .....: Retrieves a list of global groups to which a specified user belongs.
; Link .......: https://docs.microsoft.com/en-us/windows/win32/api/lmaccess/nf-lmaccess-netusergetgroups
; ===============================================================================================================================

NetUserGetGroups(UserName, ServerName := "127.0.0.1")
{
	static NERR_SUCCESS := 0
	static GROUP_USERS_INFO_0 := 0
	static MAX_PREFERRED_LENGTH := -1

	NET_API_STATUS := DllCall("netapi32\NetUserGetGroups", "wstr",  ServerName
	                                                     , "wstr",  UserName
	                                                     , "uint",  GROUP_USERS_INFO_0
	                                                     , "ptr*",  buf
	                                                     , "uint",  MAX_PREFERRED_LENGTH
	                                                     , "uint*", EntriesRead
	                                                     , "uint*", TotalEntries
	                                                     , "uint")

	if (NET_API_STATUS = NERR_SUCCESS)
	{
		addr := buf, GROUP_USERS_INFO := []
		loop % EntriesRead
		{
			GROUP_USERS_INFO.Push(StrGet(NumGet(addr + A_PtrSize * 0, "uptr"), "utf-16"))
			addr += A_PtrSize
		}

		DllCall("netapi32\NetApiBufferFree", "ptr", buf, "uint")
		return GROUP_USERS_INFO
	}

	DllCall("netapi32\NetApiBufferFree", "ptr", buf, "uint")
	return false
}

; ===============================================================================================================================

for k, v in NetUserGetGroups(A_UserName, "DC01")
	output .= k ": " v "`n"
MsgBox % output