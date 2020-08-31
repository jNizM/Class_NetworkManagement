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


	NetGroupGetInfo(GroupName, ServerName := "127.0.0.1")
	{
		static GROUP_INFO_1 := 1

		NET_API_STATUS := DllCall("netapi32\NetGroupGetInfo", "wstr", ServerName
		                                                    , "wstr", GroupName
		                                                    , "uint", GROUP_INFO_1
		                                                    , "ptr*", buf
		                                                    , "uint")

		if (NET_API_STATUS = this.NERR_SUCCESS)
		{
			GROUP_INFO := []
			GROUP_INFO["name"]    := StrGet(NumGet(buf + A_PtrSize * 0, "uptr"), "utf-16")
			GROUP_INFO["comment"] := StrGet(NumGet(buf + A_PtrSize * 1, "uptr"), "utf-16")

			this.NetApiBufferFree(buf)
			return GROUP_INFO
		}

		this.NetApiBufferFree(buf)
		return false
	}


	NetGroupGetUsers(GroupName, ServerName := "127.0.0.1")
	{
		static GROUP_USERS_INFO_0 := 0

		NET_API_STATUS := DllCall("netapi32\NetGroupGetUsers", "wstr",  ServerName
		                                                     , "wstr",  GroupName
		                                                     , "uint",  GROUP_USERS_INFO_0
		                                                     , "ptr*",  buf
		                                                     , "uint",  this.MAX_PREFERRED_LENGTH
		                                                     , "uint*", EntriesRead
		                                                     , "uint*", TotalEntries
		                                                     , "uint*", 0
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


	NetLocalGroupEnum(ServerName := "127.0.0.1")
	{
		static LOCALGROUP_INFO_1 := 1

		NET_API_STATUS := DllCall("netapi32\NetLocalGroupEnum", "wstr",  ServerName
		                                                      , "uint",  LOCALGROUP_INFO_1
		                                                      , "ptr*",  buf
		                                                      , "uint",  this.MAX_PREFERRED_LENGTH
		                                                      , "uint*", EntriesRead
		                                                      , "uint*", TotalEntries
		                                                      , "uint*", 0
		                                                      , "uint")

		if (NET_API_STATUS = this.NERR_SUCCESS)
		{
			addr := buf, LOCALGROUP_INFO := []
			loop % EntriesRead
			{
				LOCALGROUP_INFO[A_Index, "name"]    := StrGet(NumGet(addr + A_PtrSize * 0, "uptr"), "utf-16")
				LOCALGROUP_INFO[A_Index, "comment"] := StrGet(NumGet(addr + A_PtrSize * 1, "uptr"), "utf-16")
				addr += A_PtrSize * 2
			}

			this.NetApiBufferFree(buf)
			return LOCALGROUP_INFO
		}

		this.NetApiBufferFree(buf)
		return false
	}


	NetLocalGroupGetInfo(GroupName, ServerName := "127.0.0.1")
	{
		static LOCALGROUP_INFO_1 := 1

		NET_API_STATUS := DllCall("netapi32\NetLocalGroupGetInfo", "wstr", ServerName
		                                                         , "wstr", GroupName
		                                                         , "uint", LOCALGROUP_INFO_1
		                                                         , "ptr*", buf
		                                                         , "uint")

		if (NET_API_STATUS = this.NERR_SUCCESS)
		{
			LOCALGROUP_INFO := []
			LOCALGROUP_INFO["name"]    := StrGet(NumGet(buf + A_PtrSize * 0, "uptr"), "utf-16")
			LOCALGROUP_INFO["comment"] := StrGet(NumGet(buf + A_PtrSize * 1, "uptr"), "utf-16")

			this.NetApiBufferFree(buf)
			return LOCALGROUP_INFO
		}

		this.NetApiBufferFree(buf)
		return false
	}


	NetLocalGroupGetMembers(GroupName, ServerName := "127.0.0.1")
	{
		static LOCALGROUP_MEMBERS_INFO_3 := 3

		NET_API_STATUS := DllCall("netapi32\NetLocalGroupGetMembers", "wstr",  ServerName
		                                                            , "wstr",  GroupName
		                                                            , "uint",  LOCALGROUP_MEMBERS_INFO_3
		                                                            , "ptr*",  buf
		                                                            , "uint",  this.MAX_PREFERRED_LENGTH
		                                                            , "uint*", EntriesRead
		                                                            , "uint*", TotalEntries
		                                                            , "uint*", 0
		                                                            , "uint")

		if (NET_API_STATUS = this.NERR_SUCCESS)
		{
			addr := buf, LOCALGROUP_MEMBERS_INFO := []
			loop % EntriesRead
			{
				LOCALGROUP_MEMBERS_INFO.Push(StrGet(NumGet(addr + A_PtrSize * 0, "uptr"), "utf-16"))
				addr += A_PtrSize
			}

			this.NetApiBufferFree(buf)
			return LOCALGROUP_MEMBERS_INFO
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


	NetUserGetInfo(UserName, ServerName := "127.0.0.1")
	{
		static USER_INFO_4 := 4
		static TIMEQ_FOREVER := 0xFFFFFFFF
		static USER_MAXSTORAGE_UNLIMITED := 0xFFFFFFFF
		static USER_PRIV := { 0: "GUEST", 1: "USER", 2: "ADMIN" }
		static AF_OP_FLAGS := { 0: "NONE", 0x1: "PRINT", 0x2: "COMM", 0x4: "SERVER", 0x8: "ACCOUNTS" }
		static UF_FLAGS := { 0x00000001: "SCRIPT"
		                   , 0x00000002: "ACCOUNTDISABLE"
		                   , 0x00000008: "HOMEDIR_REQUIRED"
		                   , 0x00000010: "LOCKOUT"
		                   , 0x00000020: "PASSWD_NOTREQD"
		                   , 0x00000040: "PASSWD_CANT_CHANGE"
		                   , 0x00000080: "ENCRYPTED_TEXT_PASSWORD_ALLOWED"
		                   , 0x00000100: "TEMP_DUPLICATE_ACCOUNT"
		                   , 0x00000200: "NORMAL_ACCOUNT"
		                   , 0x00000800: "INTERDOMAIN_TRUST_ACCOUNT"
		                   , 0x00001000: "WORKSTATION_TRUST_ACCOUNT"
		                   , 0x00002000: "SERVER_TRUST_ACCOUNT"
		                   , 0x00010000: "DONT_EXPIRE_PASSWD"
		                   , 0x00020000: "MNS_LOGON_ACCOUNT"
		                   , 0x00040000: "SMARTCARD_REQUIRED"
		                   , 0x00080000: "TRUSTED_FOR_DELEGATION"
		                   , 0x00100000: "NOT_DELEGATED"
		                   , 0x00200000: "USE_DES_KEY_ONLY"
		                   , 0x00400000: "DONT_REQUIRE_PREAUTH"
		                   , 0x00800000: "PASSWORD_EXPIRED"
		                   , 0x01000000: "TRUSTED_TO_AUTHENTICATE_FOR_DELEGATION"
		                   , 0x02000000: "NO_AUTH_DATA_REQUIRED"
		                   , 0x04000000: "PARTIAL_SECRETS_ACCOUNT"
		                   , 0x08000000: "USE_AES_KEYS" }

		NET_API_STATUS := DllCall("netapi32\NetUserGetInfo", "wstr", ServerName
		                                                   , "wstr", UserName
		                                                   , "uint", USER_INFO_4
		                                                   , "ptr*", buf
		                                                   , "uint")

		if (NET_API_STATUS = this.NERR_SUCCESS)
		{
			USER_INFO := []
			USER_INFO["name"]             := StrGet(NumGet(buf + A_PtrSize * 0, "uptr"), "utf-16")
			; USER_INFO["password"]       := StrGet(NumGet(buf + A_PtrSize * 1, "uptr"), "utf-16")   ; returns a NULL pointer
			USER_INFO["password_age"]     := this.ConvertAge(NumGet(buf + A_PtrSize * 2, "uint"))
			USER_INFO["priv"]             := USER_PRIV[NumGet(buf + A_PtrSize * 2 + 4, "uint")]
			USER_INFO["home_dir"]         := StrGet(NumGet(buf + A_PtrSize * 3, "uptr"), "utf-16")
			USER_INFO["comment"]          := StrGet(NumGet(buf + A_PtrSize * 4, "uptr"), "utf-16")
			USER_INFO["flags"]            := flags := Format("{:#x}", NumGet(buf + A_PtrSize * 5, "uint"))
			for k, v in UF_FLAGS
				if (k & flags)
					flags_list .= v " | "
			USER_INFO["flags_list"]       := SubStr(flags_list, 1, -3)
			USER_INFO["script_path"]      := StrGet(NumGet(buf + A_PtrSize * 6, "uptr"), "utf-16")
			USER_INFO["auth_flags"]       := auth_flags := Format("{:#x}", NumGet(buf + A_PtrSize * 7, "uint"))
			for k, v in AF_OP_FLAGS
				if (k & auth_flags)
					auth_flags_list .= v " | "
			USER_INFO["auth_flags_list"]  := SubStr(auth_flags_list, 1, -3)
			USER_INFO["full_name"]        := StrGet(NumGet(buf + A_PtrSize * 8, "uptr"), "utf-16")
			USER_INFO["usr_comment"]      := StrGet(NumGet(buf + A_PtrSize * 9, "uptr"), "utf-16")
			USER_INFO["parms"]            := StrGet(NumGet(buf + A_PtrSize * 10, "uptr"), "utf-16")
			USER_INFO["workstations"]     := StrGet(NumGet(buf + A_PtrSize * 11, "uptr"), "utf-16")
			USER_INFO["last_logon"]       := this.ConvertUnixTime(NumGet(buf + A_PtrSize * 12, "uint"))
			; USER_INFO["last_logoff"]    := NumGet(buf + A_PtrSize * 12 + 4, "uint")   ; This member is currently not used.
			USER_INFO["acct_expires"]     := (NumGet(buf + A_PtrSize * 13, "uint") = TIMEQ_FOREVER)
			                               ? "NEVER"
			                               : this.ConvertUnixTime(NumGet(buf + A_PtrSize * 13, "uint"))
			USER_INFO["max_storage"]      := (NumGet(buf + A_PtrSize * 13 + 4, "uint") = USER_MAXSTORAGE_UNLIMITED)
			                               ? "UNLIMITED"
			                               : NumGet(buf + A_PtrSize * 13 + 4, "uint")
			USER_INFO["units_per_week"]   := NumGet(buf + A_PtrSize * 14, "uint")
			USER_INFO["logon_hours"]      := NumGet(buf + A_PtrSize * 15, "uchar*")   ; todo
			USER_INFO["bad_pw_count"]     := NumGet(buf + A_PtrSize * 16, "uint")
			USER_INFO["num_logons"]       := NumGet(buf + A_PtrSize * 16 + 4, "uint")
			USER_INFO["logon_server"]     := StrGet(NumGet(buf + A_PtrSize * 17, "uptr"), "utf-16")
			USER_INFO["country_code"]     := NumGet(buf + A_PtrSize * 18, "uint")
			USER_INFO["code_page"]        := NumGet(buf + A_PtrSize * 18 + 4, "uint")
			USER_INFO["user_sid"]         := PSID := NumGet(buf + A_PtrSize * 19, "ptr")
			USER_INFO["user_sid_string"]  := this.ConvertSidToStringSid(PSID)
			USER_INFO["primary_group_id"] := NumGet(buf + A_PtrSize * 19 + 4, "uint")
			USER_INFO["profile"]          := StrGet(NumGet(buf + A_PtrSize * 20, "uptr"), "utf-16")
			USER_INFO["home_dir_drive"]   := StrGet(NumGet(buf + A_PtrSize * 21, "uptr"), "utf-16")
			USER_INFO["password_expired"] := NumGet(buf + A_PtrSize * 22, "uint")

			this.NetApiBufferFree(buf)
			return USER_INFO
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

	ConvertAge(value)
	{
		age := A_Now
		age += -(value), s
		FormatTime, age, % age, yyyy-MM-dd HH:mm:ss
		return age
	}


	ConvertSidToStringSid(PSID)
	{
		DllCall("advapi32\ConvertSidToStringSid", "ptr", PSID, "ptr*", StringSid)
		VarSetCapacity(SID, DllCall("lstrlenW", "ptr", StringSid) * 2)
		DllCall("lstrcpyW", "str", SID, "ptr", StringSid)
		DllCall("LocalFree", "ptr", StringSid)
		return SID
	}


	ConvertUnixTime(value)
	{
		VarSetCapacity(TIME_ZONE_INFORMATION, 44 + (64 << !!A_IsUnicode), 0)
		TIME_ZONE_ID := DllCall("GetTimeZoneInformation", "ptr", &TIME_ZONE_INFORMATION, "uint")

		unix := 1970
		unix += value, s
		unix += (TIME_ZONE_ID = 1 ? 1 : 0), hours
		FormatTime, unix, % unix, yyyy-MM-dd HH:mm:ss
		return unix
	}

	NetApiBufferFree(buffer)
	{
		NET_API_STATUS := DllCall("netapi32\NetApiBufferFree", "ptr", buffer, "uint")

		if (NET_API_STATUS = this.NERR_SUCCESS)
			return true
		return false
	}

}

; ===============================================================================================================================