; ===============================================================================================================================
; Function ...: NetLocalGroupGetInfo
; Return .....: Retrieves information about a particular local group account on a server.
; Link .......: https://docs.microsoft.com/en-us/windows/win32/api/lmaccess/nf-lmaccess-netlocalgroupgetinfo
; ===============================================================================================================================

NetLocalGroupGetInfo(GroupName, ServerName := "127.0.0.1")
{
	static NERR_SUCCESS := 0
	static LOCALGROUP_INFO_1 := 1

	NET_API_STATUS := DllCall("netapi32\NetLocalGroupGetInfo", "wstr", ServerName
	                                                         , "wstr", GroupName
	                                                         , "uint", LOCALGROUP_INFO_1
	                                                         , "ptr*", buf
	                                                         , "uint")

	if (NET_API_STATUS = NERR_SUCCESS)
	{
		LOCALGROUP_INFO := []
		LOCALGROUP_INFO["name"]    := StrGet(NumGet(buf + A_PtrSize * 0, "uptr"), "utf-16")
		LOCALGROUP_INFO["comment"] := StrGet(NumGet(buf + A_PtrSize * 1, "uptr"), "utf-16")

		DllCall("netapi32\NetApiBufferFree", "ptr", buf, "uint")
		return LOCALGROUP_INFO
	}

	DllCall("netapi32\NetApiBufferFree", "ptr", buf, "uint")
	return false
}

; ===============================================================================================================================

for k, v in NetLocalGroupGetInfo("L_GROUP_TEST", "DC01")
	output .= k ": " v "`n"
MsgBox % output