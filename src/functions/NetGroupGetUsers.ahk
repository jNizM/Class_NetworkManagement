; ===============================================================================================================================
; Function ...: NetGroupGetUsers
; Return .....: Retrieves a list of the members in a particular global group in the security database, which is the security
;               accounts manager (SAM) database or, in the case of domain controllers, the Active Directory.
; Link .......: https://docs.microsoft.com/en-us/windows/win32/api/lmaccess/nf-lmaccess-netgroupgetusers
; ===============================================================================================================================

NetGroupGetUsers(GroupName, ServerName := "127.0.0.1")
{
	static NERR_SUCCESS := 0
	static GROUP_USERS_INFO_0 := 0
	static MAX_PREFERRED_LENGTH := -1

	NET_API_STATUS := DllCall("netapi32\NetGroupGetUsers", "wstr",  ServerName
	                                                     , "wstr",  GroupName
	                                                     , "uint",  GROUP_USERS_INFO_0
	                                                     , "ptr*",  buf
	                                                     , "uint",  MAX_PREFERRED_LENGTH
	                                                     , "uint*", EntriesRead
	                                                     , "uint*", TotalEntries
	                                                     , "uint*", 0
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

for k, v in NetGroupGetUsers("G_GROUP_TEST", "DC01")
	output .= k ": " v "`n"
MsgBox % output