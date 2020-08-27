; ===============================================================================================================================
; Function ...: NetServerGetInfo
; Return .....: Retrieves current configuration information for the specified server.
; Link .......: https://docs.microsoft.com/en-us/windows/win32/api/lmserver/nf-lmserver-netservergetinfo
; ===============================================================================================================================

NetServerGetInfo(ServerName := "127.0.0.1")
{
	static NERR_SUCCESS := 0
	static SERVER_INFO_101 := 101
	static PLATFORM_ID := { 300: "DOS", 400: "OS2", 500: "NT", 600: "OSF", 700: "VMS" }
	static SV_TYPE := { 0x00000001: "WORKSTATION"
	                  , 0x00000002: "SERVER"
	                  , 0x00000004: "SQLSERVER"
	                  , 0x00000008: "DOMAIN_CTRL"
	                  , 0x00000010: "DOMAIN_BAKCTRL"
	                  , 0x00000020: "TIME_SOURCE"
	                  , 0x00000040: "AFP"
	                  , 0x00000080: "NOVELL"
	                  , 0x00000100: "DOMAIN_MEMBER"
	                  , 0x00000200: "PRINTQ_SERVER"
	                  , 0x00000400: "DIALIN_SERVER"
	                  , 0x00000800: "XENIX_SERVER"
	                  , 0x00001000: "NT"
	                  , 0x00002000: "WFW"
	                  , 0x00004000: "SERVER_MFPN"
	                  , 0x00008000: "SERVER_NT"
	                  , 0x00010000: "POTENTIAL_BROWSER"
	                  , 0x00020000: "BACKUP_BROWSER"
	                  , 0x00040000: "MASTER_BROWSER"
	                  , 0x00080000: "DOMAIN_MASTER"
	                  , 0x00100000: "SERVER_OSF"
	                  , 0x00200000: "SERVER_VMS"
	                  , 0x00400000: "WINDOWS"
	                  , 0x00800000: "DFS"
	                  , 0x01000000: "CLUSTER_NT"
	                  , 0x02000000: "TERMINALSERVER"
	                  , 0x04000000: "CLUSTER_VS_NT"
	                  , 0x10000000: "DCE"
	                  , 0x20000000: "ALTERNATE_XPORT"
	                  , 0x40000000: "LOCAL_LIST_ONLY"
	                  , 0x80000000: "DOMAIN_ENUM" }

	NET_API_STATUS := DllCall("netapi32\NetServerGetInfo", "wstr", ServerName
	                                                     , "uint", SERVER_INFO_101
	                                                     , "ptr*", buf
	                                                     , "uint")

	if (NET_API_STATUS = NERR_SUCCESS)
	{
		SERVER_INFO := []
		SERVER_INFO["platform_id"]   := PLATFORM_ID[NumGet(buf + 0, "uint")]
		SERVER_INFO["name"]          := StrGet(NumGet(buf + A_PtrSize * 1, "uptr"), "utf-16")
		SERVER_INFO["version_major"] := NumGet(buf + A_PtrSize * 2, "uint")
		SERVER_INFO["version_minor"] := NumGet(buf + A_PtrSize * 2 + 4, "uint")
		SERVER_INFO["type"]          := types := Format("{:#x}", NumGet(buf + A_PtrSize * 3, "uint"))
		for k, v in SV_TYPE
			if (k & types)
				type_list .= v " | "
		SERVER_INFO["type_list"]     := SubStr(type_list, 1, -3)
		SERVER_INFO["comment"]       := StrGet(NumGet(buf + A_PtrSize * 4, "uptr"), "utf-16")

		DllCall("netapi32\NetApiBufferFree", "ptr", buf, "uint")
		return SERVER_INFO
	}

	DllCall("netapi32\NetApiBufferFree", "ptr", buf, "uint")
	return false
}

; ===============================================================================================================================

for k, v in NetServerGetInfo()
	output .= k ": " v "`n"
MsgBox % output
output := ""