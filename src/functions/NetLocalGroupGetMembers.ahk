; ===============================================================================================================================
; Function ...: NetLocalGroupGetMembers
; Return .....: Retrieves a list of the members of a particular local group in the security database, which is the security
;               accounts manager (SAM) database or, in the case of domain controllers, the Active Directory. Local group members
;               can be users or global groups.
; Link .......: https://docs.microsoft.com/en-us/windows/win32/api/lmaccess/nf-lmaccess-netlocalgroupgetmembers
; ===============================================================================================================================

NetLocalGroupGetMembers(GroupName, ServerName := "127.0.0.1")
{
	static NERR_SUCCESS := 0
	static LOCALGROUP_MEMBERS_INFO_3 := 3
	static MAX_PREFERRED_LENGTH := -1

	NET_API_STATUS := DllCall("netapi32\NetLocalGroupGetMembers", "wstr",  ServerName
	                                                            , "wstr",  GroupName
	                                                            , "uint",  LOCALGROUP_MEMBERS_INFO_3
	                                                            , "ptr*",  buf
	                                                            , "uint",  MAX_PREFERRED_LENGTH
	                                                            , "uint*", EntriesRead
	                                                            , "uint*", TotalEntries
	                                                            , "uint*", 0
	                                                            , "uint")

	if (NET_API_STATUS = NERR_SUCCESS)
	{
		addr := buf, LOCALGROUP_MEMBERS_INFO := []
		loop % EntriesRead
		{
			LOCALGROUP_MEMBERS_INFO.Push(StrGet(NumGet(addr + A_PtrSize * 0, "uptr"), "utf-16"))
			addr += A_PtrSize
		}

		DllCall("netapi32\NetApiBufferFree", "ptr", buf, "uint")
		return LOCALGROUP_MEMBERS_INFO
	}

	DllCall("netapi32\NetApiBufferFree", "ptr", buf, "uint")
	return false
}

; ===============================================================================================================================

for k, v in NetLocalGroupGetMembers("L_GROUP_TEST", "DC01")
	output .= k ": " v "`n"
MsgBox % output