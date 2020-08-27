; ===============================================================================================================================
; Function ...: NetWkstaUserGetInfo
; Return .....: Information about the currently logged-on user.
; Link .......: https://docs.microsoft.com/en-us/windows/win32/api/lmwksta/nf-lmwksta-netwkstausergetinfo
; ===============================================================================================================================

NetWkstaUserGetInfo()
{
	static NERR_SUCCESS := 0
	static WKSTA_USER_INFO_1 := 1

	NET_API_STATUS := DllCall("netapi32\NetWkstaUserGetInfo", "ptr",  0
	                                                        , "uint", WKSTA_USER_INFO_1
	                                                        , "ptr*", buf
	                                                        , "uint")

	if (NET_API_STATUS = NERR_SUCCESS)
	{
		WKSTA_USER_INFO := []
		WKSTA_USER_INFO["username"]     := StrGet(NumGet(buf + A_PtrSize * 0, "uptr"), "utf-16")
		WKSTA_USER_INFO["logon_domain"] := StrGet(NumGet(buf + A_PtrSize * 1, "uptr"), "utf-16")
		WKSTA_USER_INFO["oth_domains"]  := StrGet(NumGet(buf + A_PtrSize * 2, "uptr"), "utf-16")
		WKSTA_USER_INFO["logon_server"] := StrGet(NumGet(buf + A_PtrSize * 3, "uptr"), "utf-16")

		DllCall("netapi32\NetApiBufferFree", "ptr", buf, "uint")
		return WKSTA_USER_INFO
	}

	DllCall("netapi32\NetApiBufferFree", "ptr", buf, "uint")
	return false
}

; ===============================================================================================================================

for k, v in NetWkstaUserGetInfo()
	output .= k ": " v "`n"
MsgBox % output