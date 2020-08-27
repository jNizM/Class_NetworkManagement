; ===============================================================================================================================
; Function ...: NetWkstaGetInfo
; Return .....: Information about the configuration of a workstation.
; Link .......: https://docs.microsoft.com/en-us/windows/win32/api/lmwksta/nf-lmwksta-netwkstagetinfo
; ===============================================================================================================================

NetWkstaGetInfo(ServerName := "127.0.0.1")
{
	static NERR_SUCCESS := 0
	static WKSTA_INFO_102 := 102
	static PLATFORM_ID := { 300: "DOS", 400: "OS2", 500: "NT", 600: "OSF", 700: "VMS" }

	NET_API_STATUS := DllCall("netapi32\NetWkstaGetInfo", "wstr", ServerName
	                                                    , "uint", WKSTA_INFO_102
	                                                    , "ptr*", buf
	                                                    , "uint")

	if (NET_API_STATUS = NERR_SUCCESS)
	{
		WKSTA_INFO := []
		WKSTA_INFO["platform_id"]     := PLATFORM_ID[NumGet(buf + 0, "uint")]
		WKSTA_INFO["computername"]    := StrGet(NumGet(buf + A_PtrSize * 1, "uptr"), "utf-16")
		WKSTA_INFO["langroup"]        := StrGet(NumGet(buf + A_PtrSize * 2, "uptr"), "utf-16")
		WKSTA_INFO["ver_major"]       := NumGet(buf + A_PtrSize * 3, "uint")
		WKSTA_INFO["ver_minor"]       := NumGet(buf + A_PtrSize * 3 + 4, "uint")
		WKSTA_INFO["lanroot"]         := StrGet(NumGet(buf + A_PtrSize * 4, "uptr"), "utf-16")
		WKSTA_INFO["logged_on_users"] := NumGet(buf + A_PtrSize * 5, "uint")

		DllCall("netapi32\NetApiBufferFree", "ptr", buf, "uint")
		return WKSTA_INFO
	}

	DllCall("netapi32\NetApiBufferFree", "ptr", buf, "uint")
	return false
}

; ===============================================================================================================================

for k, v in NetWkstaGetInfo()
	output .= k ": " v "`n"
MsgBox % output