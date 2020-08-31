; ===============================================================================================================================
; Function ...: NetGroupGetInfo
; Return .....: Retrieves information about a particular global group in the security database, which is the security accounts
;               manager (SAM) database or, in the case of domain controllers, the Active Directory.
; Link .......: https://docs.microsoft.com/en-us/windows/win32/api/lmaccess/nf-lmaccess-netgroupgetinfo
; ===============================================================================================================================

NetGroupGetInfo(GroupName, ServerName := "127.0.0.1")
{
	static NERR_SUCCESS := 0
	static GROUP_INFO_1 := 1

	NET_API_STATUS := DllCall("netapi32\NetGroupGetInfo", "wstr", ServerName
	                                                    , "wstr", GroupName
	                                                    , "uint", GROUP_INFO_1
	                                                    , "ptr*", buf
	                                                    , "uint")

	if (NET_API_STATUS = NERR_SUCCESS)
	{
		GROUP_INFO := []
		GROUP_INFO["name"]    := StrGet(NumGet(buf + A_PtrSize * 0, "uptr"), "utf-16")
		GROUP_INFO["comment"] := StrGet(NumGet(buf + A_PtrSize * 1, "uptr"), "utf-16")

		DllCall("netapi32\NetApiBufferFree", "ptr", buf, "uint")
		return GROUP_INFO
	}

	DllCall("netapi32\NetApiBufferFree", "ptr", buf, "uint")
	return false
}

; ===============================================================================================================================

for k, v in NetGroupGetInfo("G_GROUP_TEST", "DC01")
	output .= k ": " v "`n"
MsgBox % output