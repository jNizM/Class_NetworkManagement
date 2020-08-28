; ===============================================================================================================================
; AutoHotkey wrapper for Network Management technology
;
; Author ....: jNizM
; Released ..: 2020-08-22
; Modified ..: 2020-08-28
; Github ....: https://github.com/jNizM/Class_NetworkManagement
; Forum .....: https://www.autohotkey.com/boards/viewtopic.php?f=6&t=80382
; ===============================================================================================================================


class NetworkManagement
{
	static NERR_SUCCESS := 0
	static ERROR_MORE_DATA := 234
	static MAX_PREFERRED_LENGTH := -1


	; ===== PUBLIC METHODS ======================================================================================================

	NetGetAnyDCName(ServerName := "", DomainName := "")
	{
		NET_API_STATUS := DllCall("netapi32\NetGetAnyDCName", "wstr", ServerName
		                                                    , "wstr", DomainName
		                                                    , "ptr*", buf
		                                                    , "uint")

		if (NET_API_STATUS = this.NERR_SUCCESS)
		{
			DomainController := buf
			this.NetApiBufferFree(buf)
			return StrGet(DomainController)
		}
		else
		{
			this.NetApiBufferFree(buf)
			switch NET_API_STATUS
			{
				case 0x00000057:
					return "ERROR_INVALID_PARAMETER"
				case 0x0000054B:
					return "ERROR_NO_SUCH_DOMAIN"
				case 0x00000032:
					return "ERROR_NOT_SUPPORTED"
				case 0x00000035:
					return "ERROR_BAD_NETPATH"
				case 0x000004BA:
					return "ERROR_INVALID_COMPUTERNAME"
				case 0x00002558:
					return "DNS_ERROR_INVALID_NAME_CHAR"
				case 0x00002554:
					return "DNS_ERROR_NON_RFC_NAME"
				case 0x0000007B:
					return "ERROR_INVALID_NAME"
				case 0x00000995:
					return "NERR_DCNotFound"
				case 0x0000085A:
					return "NERR_WkstaNotStarted"
				case 0x000006BA:
					return "RPC_S_SERVER_UNAVAILABLE"
				case 0x8001011C:
					return "RPC_E_REMOTE_DISABLED"
				default:
					return "Other error"
			}
		}
	}


	NetGetDCName(ServerName := "", DomainName := "")
	{
		NET_API_STATUS := DllCall("netapi32\NetGetDCName", "wstr", ServerName
		                                                 , "wstr", DomainName
		                                                 , "ptr*", buf
		                                                 , "uint")

		if (NET_API_STATUS = this.NERR_SUCCESS)
		{
			DomainController := buf
			this.NetApiBufferFree(buf)
			return StrGet(DomainController)
		}
		else
		{
			this.NetApiBufferFree(buf)
			switch NET_API_STATUS
			{
				case 0x00000057:
					return "ERROR_INVALID_PARAMETER"
				case 0x0000054B:
					return "ERROR_NO_SUCH_DOMAIN"
				case 0x00000032:
					return "ERROR_NOT_SUPPORTED"
				case 0x00000035:
					return "ERROR_BAD_NETPATH"
				case 0x000004BA:
					return "ERROR_INVALID_COMPUTERNAME"
				case 0x00002558:
					return "DNS_ERROR_INVALID_NAME_CHAR"
				case 0x00002554:
					return "DNS_ERROR_NON_RFC_NAME"
				case 0x0000007B:
					return "ERROR_INVALID_NAME"
				case 0x00000995:
					return "NERR_DCNotFound"
				case 0x0000085A:
					return "NERR_WkstaNotStarted"
				case 0x000006BA:
					return "RPC_S_SERVER_UNAVAILABLE"
				case 0x8001011C:
					return "RPC_E_REMOTE_DISABLED"
				default:
					return "Other error"
			}
		}
	}


	NetGetJoinInformation(Server := "127.0.0.1")
	{
		static JOIN_STATUS := { 0: "Unknown", 1: "Unjoined", 2: "Workgroup", 3: "Domain" }

		NET_API_STATUS := DllCall("netapi32\NetGetJoinInformation", "wstr", Server
		                                                          , "ptr*", BufferName
		                                                          , "int*", BufferType
		                                                          , "uint")

		if (NET_API_STATUS = this.NERR_SUCCESS)
		{
			JOIN_INFO := []
			JOIN_INFO["Name"] := StrGet(BufferName)
			JOIN_INFO["Type"] := JOIN_STATUS[BufferType]

			this.NetApiBufferFree(buf)
			return JOIN_INFO
		}

		this.NetApiBufferFree(buf)
		return false
	}


	NetGroupEnum(ServerName := "127.0.0.1")
	{
		static GROUP_INFO_1 := 1

		NET_API_STATUS := DllCall("netapi32\NetGroupEnum", "wstr",  ServerName
		                                                 , "uint",  GROUP_INFO_1
		                                                 , "ptr*",  buf
		                                                 , "uint",  this.MAX_PREFERRED_LENGTH
		                                                 , "uint*", EntriesRead
		                                                 , "uint*", TotalEntries
		                                                 , "uint*", 0
		                                                 , "uint")

		if (NET_API_STATUS = this.NERR_SUCCESS)
		{
			addr := buf, GROUP_INFO := []
			loop % EntriesRead
			{
				GROUP_INFO[A_Index, "name"]    := StrGet(NumGet(addr + A_PtrSize * 0, "uptr"), "utf-16")
				GROUP_INFO[A_Index, "comment"] := StrGet(NumGet(addr + A_PtrSize * 1, "uptr"), "utf-16")
				addr += A_PtrSize * 2
			}

			this.NetApiBufferFree(buf)
			return GROUP_INFO
		}

		this.NetApiBufferFree(buf)
		return false
	}


	NetServerGetInfo(ServerName := "127.0.0.1")
	{
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

		if (NET_API_STATUS = this.NERR_SUCCESS)
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
			this.NetApiBufferFree(buf)
			return SERVER_INFO
		}

		this.NetApiBufferFree(buf)
		return false
	}


	NetServerTransportEnum(ServerName := "127.0.0.1")
	{
		static SERVER_TRANSPORT_INFO_1 := 1

		NET_API_STATUS := DllCall("netapi32\NetServerTransportEnum", "wstr",  ServerName
		                                                           , "uint",  SERVER_TRANSPORT_INFO_1
		                                                           , "ptr*",  buf
		                                                           , "uint",  this.MAX_PREFERRED_LENGTH
		                                                           , "uint*", EntriesRead
		                                                           , "uint*", TotalEntries
		                                                           , "uint*", 0
		                                                           , "uint")

		if ((NET_API_STATUS = this.NERR_SUCCESS) || (NET_API_STATUS = this.ERROR_MORE_DATA))
		{
			addr := buf, SERVER_TRANSPORT_INFO := []
			loop % EntriesRead
			{
				SERVER_TRANSPORT_INFO[A_Index, "numberofvcs"]            := NumGet(addr + 0, "uint")
				SERVER_TRANSPORT_INFO[A_Index, "transportname"]          := StrGet(NumGet(addr + A_PtrSize * 1, "uptr"), "utf-16")
				SERVER_TRANSPORT_INFO[A_Index, "transportaddress"]       := NumGet(addr + A_PtrSize * 2, "ptr")							; todo
				SERVER_TRANSPORT_INFO[A_Index, "transportaddresslength"] := NumGet(addr + A_PtrSize * 3, "uint")
				SERVER_TRANSPORT_INFO[A_Index, "networkaddress"]         := StrGet(NumGet(addr + A_PtrSize * 4, "uptr"), "utf-16")
				SERVER_TRANSPORT_INFO[A_Index, "domain"]                 := StrGet(NumGet(addr + A_PtrSize * 5, "uptr"), "utf-16")
				addr += A_PtrSize * 6
			}

			this.NetApiBufferFree(buf)
			return SERVER_TRANSPORT_INFO
		}

		this.NetApiBufferFree(buf)
		return false
	}


	NetUserGetGroups(UserName, ServerName := "127.0.0.1")
	{
		static GROUP_USERS_INFO_0 := 0

		NET_API_STATUS := DllCall("netapi32\NetUserGetGroups", "wstr",  ServerName
		                                                     , "wstr",  UserName
		                                                     , "uint",  GROUP_USERS_INFO_0
		                                                     , "ptr*",  buf
		                                                     , "uint",  this.MAX_PREFERRED_LENGTH
		                                                     , "uint*", EntriesRead
		                                                     , "uint*", TotalEntries
		                                                     , "uint")

		if (NET_API_STATUS = this.NERR_SUCCESS)
		{
			addr := buf, GROUP_USERS_INFO := []
			loop % EntriesRead
			{
				GROUP_USERS_INFO.Push(StrGet(NumGet(addr + A_PtrSize * 0, "uptr"), "utf-16"))
				addr += A_PtrSize
			}

			this.NetApiBufferFree(buf)
			return GROUP_USERS_INFO
		}

		this.NetApiBufferFree(buf)
		return false
	}


	NetUserGetLocalGroups(UserName, Domain := "", ServerName := "127.0.0.1")
	{
		static LOCALGROUP_USERS_INFO_0 := 0
		static LG_INCLUDE_INDIRECT := 0x0001

		NET_API_STATUS := DllCall("netapi32\NetUserGetLocalGroups", "wstr",  ServerName
		                                                          , "wstr",  Domain . UserName
		                                                          , "uint",  LOCALGROUP_USERS_INFO_0
		                                                          , "uint",  LG_INCLUDE_INDIRECT
		                                                          , "ptr*",  buf
		                                                          , "uint",  this.MAX_PREFERRED_LENGTH
		                                                          , "uint*", EntriesRead
		                                                          , "uint*"  TotalEntries
		                                                          , "uint")

		if (NET_API_STATUS = this.NERR_SUCCESS)
		{
			addr := buf, LOCALGROUP_USERS_INFO := []
			loop % EntriesRead
			{
				LOCALGROUP_USERS_INFO.Push(StrGet(NumGet(addr + A_PtrSize * 0, "uptr"), "utf-16"))
				addr += A_PtrSize
			}

			this.NetApiBufferFree(buf)
			return LOCALGROUP_USERS_INFO
		}

		this.NetApiBufferFree(buf)
		return false
	}


	NetWkstaGetInfo(ServerName := "127.0.0.1")
	{
		static WKSTA_INFO_102 := 102
		static PLATFORM_ID := { 300: "DOS", 400: "OS2", 500: "NT", 600: "OSF", 700: "VMS" }

		NET_API_STATUS := DllCall("netapi32\NetWkstaGetInfo", "wstr", ServerName
		                                                    , "uint", WKSTA_INFO_102
		                                                    , "ptr*", buf
		                                                    , "uint")

		if (NET_API_STATUS = this.NERR_SUCCESS)
		{
			WKSTA_INFO := []
			WKSTA_INFO["platform_id"]     := PLATFORM_ID[NumGet(buf + 0, "uint")]
			WKSTA_INFO["computername"]    := StrGet(NumGet(buf + A_PtrSize * 1, "uptr"), "utf-16")
			WKSTA_INFO["langroup"]        := StrGet(NumGet(buf + A_PtrSize * 2, "uptr"), "utf-16")
			WKSTA_INFO["ver_major"]       := NumGet(buf + A_PtrSize * 3, "uint")
			WKSTA_INFO["ver_minor"]       := NumGet(buf + A_PtrSize * 3 + 4, "uint")
			WKSTA_INFO["lanroot"]         := StrGet(NumGet(buf + A_PtrSize * 4, "uptr"), "utf-16")
			WKSTA_INFO["logged_on_users"] := NumGet(buf + A_PtrSize * 5, "uint")

			this.NetApiBufferFree(buf)
			return WKSTA_INFO
		}

		this.NetApiBufferFree(buf)
		return false
	}


	NetWkstaTransportEnum(ServerName := "127.0.0.1")
	{
		static WKSTA_TRANSPORT_INFO_0 := 0

		NET_API_STATUS := DllCall("netapi32\NetWkstaTransportEnum", "wstr",  ServerName
		                                                          , "uint",  WKSTA_TRANSPORT_INFO_0
		                                                          , "ptr*",  buf
		                                                          , "uint",  this.MAX_PREFERRED_LENGTH
		                                                          , "uint*", EntriesRead
		                                                          , "uint*", TotalEntries
		                                                          , "uint*", 0
		                                                          , "uint")

		if ((NET_API_STATUS = this.NERR_SUCCESS) || (NET_API_STATUS = this.ERROR_MORE_DATA))
		{
			addr := buf, WKSTA_TRANSPORT_INFO := []
			loop % EntriesRead
			{
				WKSTA_TRANSPORT_INFO[A_Index, "quality_of_service"] := NumGet(addr + 0, "uint")
				WKSTA_TRANSPORT_INFO[A_Index, "number_of_vcs"]      := NumGet(addr + 4, "uint")
				WKSTA_TRANSPORT_INFO[A_Index, "transport_name"]     := StrGet(NumGet(addr + A_PtrSize * 2, "uptr"), "utf-16")
				WKSTA_TRANSPORT_INFO[A_Index, "transport_address"]  := StrGet(NumGet(addr + A_PtrSize * 3, "uptr"), "utf-16")
				WKSTA_TRANSPORT_INFO[A_Index, "wan_ish"]            := NumGet(addr + A_PtrSize * 4, "int")
				addr += 8 + A_PtrSize * 3
			}

			this.NetApiBufferFree(buf)
			return WKSTA_TRANSPORT_INFO
		}

		this.NetApiBufferFree(buf)
		return false
	}


	NetWkstaUserEnum(ServerName := "127.0.0.1")
	{
		static WKSTA_USER_INFO_1 := 1

		NET_API_STATUS := DllCall("netapi32\NetWkstaUserEnum", "wstr",  ServerName
		                                                     , "uint",  WKSTA_USER_INFO_1
		                                                     , "ptr*",  buf
		                                                     , "uint",  this.MAX_PREFERRED_LENGTH
		                                                     , "uint*", EntriesRead
		                                                     , "uint*", TotalEntries
		                                                     , "uint*", 0
		                                                     , "uint")

		if ((NET_API_STATUS = this.NERR_SUCCESS) || (NET_API_STATUS = this.ERROR_MORE_DATA))
		{
			addr := buf, WKSTA_USER_INFO := []
			loop % EntriesRead
			{
				WKSTA_USER_INFO[A_Index, "username"]     := StrGet(NumGet(addr + A_PtrSize * 0, "uptr"), "utf-16")
				WKSTA_USER_INFO[A_Index, "logon_domain"] := StrGet(NumGet(addr + A_PtrSize * 1, "uptr"), "utf-16")
				WKSTA_USER_INFO[A_Index, "oth_domains"]  := StrGet(NumGet(addr + A_PtrSize * 2, "uptr"), "utf-16")
				WKSTA_USER_INFO[A_Index, "logon_server"] := StrGet(NumGet(addr + A_PtrSize * 3, "uptr"), "utf-16")
				addr += A_PtrSize * 4
			}

			this.NetApiBufferFree(buf)
			return WKSTA_USER_INFO
		}

		this.NetApiBufferFree(buf)
		return false
	}


	NetWkstaUserGetInfo()
	{
		static WKSTA_USER_INFO_1 := 1

		NET_API_STATUS := DllCall("netapi32\NetWkstaUserGetInfo", "ptr",  0
		                                                        , "uint", WKSTA_USER_INFO_1
		                                                        , "ptr*", buf
		                                                        , "uint")

		if (NET_API_STATUS = this.NERR_SUCCESS)
		{
			WKSTA_USER_INFO := []
			WKSTA_USER_INFO["username"]     := StrGet(NumGet(buf + A_PtrSize * 0, "uptr"), "utf-16")
			WKSTA_USER_INFO["logon_domain"] := StrGet(NumGet(buf + A_PtrSize * 1, "uptr"), "utf-16")
			WKSTA_USER_INFO["oth_domains"]  := StrGet(NumGet(buf + A_PtrSize * 2, "uptr"), "utf-16")
			WKSTA_USER_INFO["logon_server"] := StrGet(NumGet(buf + A_PtrSize * 3, "uptr"), "utf-16")

			this.NetApiBufferFree(buf)
			return WKSTA_USER_INFO
		}

		this.NetApiBufferFree(buf)
		return false
	}


	; ===== PRIVATE METHODS =====================================================================================================

	NetApiBufferFree(buffer)
	{
		NET_API_STATUS := DllCall("netapi32\NetApiBufferFree", "ptr", buffer, "uint")

		if (NET_API_STATUS = this.NERR_SUCCESS)
			return true
		return false
	}

}

; ===============================================================================================================================