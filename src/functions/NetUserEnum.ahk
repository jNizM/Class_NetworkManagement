; ===============================================================================================================================
; Function ...: NetUserEnum
; Return .....: Retrieves information about all user accounts on a server.
; Link .......: https://docs.microsoft.com/en-us/windows/win32/api/lmaccess/nf-lmaccess-netuserenum
; ===============================================================================================================================

NetUserEnum(ServerName := "127.0.0.1", Filter := 0x0002)
{
	static NERR_SUCCESS := 0
	static ERROR_MORE_DATA := 234
	static USER_INFO_3 := 3
	static MAX_PREFERRED_LENGTH := -1
	static TIMEQ_FOREVER := 0xFFFFFFFF
	static USER_MAXSTORAGE_UNLIMITED := 0xFFFFFFFF
	static USER_PRIV := { 0: "GUEST", 1: "USER", 2: "ADMIN" }
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

	NET_API_STATUS := DllCall("netapi32\NetUserEnum", "wstr",  ServerName
	                                                , "uint",  USER_INFO_3
	                                                , "uint",  Filter
	                                                , "ptr*",  buf
	                                                , "uint",  MAX_PREFERRED_LENGTH
	                                                , "uint*", EntriesRead
	                                                , "uint*", TotalEntries
	                                                , "uint*", 0
	                                                , "uint")

	if ((NET_API_STATUS = NERR_SUCCESS) || (NET_API_STATUS = ERROR_MORE_DATA))
	{
		addr := buf, USER_INFO := []
		loop % EntriesRead
		{
			USER_INFO[A_Index, "name"]             := StrGet(NumGet(addr + A_PtrSize * 0, "uptr"), "utf-16")
			;USER_INFO[A_Index, "password"]        := StrGet(NumGet(addr + A_PtrSize * 1, "uptr"), "utf-16")   ; returns a NULL pointer
			USER_INFO[A_Index, "password_age"]     := ConvertAge(NumGet(addr + A_PtrSize * 2, "uint"))
			USER_INFO[A_Index, "priv"]             := USER_PRIV[NumGet(addr + A_PtrSize * 2 + 4, "uint")]
			USER_INFO[A_Index, "home_dir"]         := StrGet(NumGet(addr + A_PtrSize * 3, "uptr"), "utf-16")
			USER_INFO[A_Index, "comment"]          := StrGet(NumGet(addr + A_PtrSize * 4, "uptr"), "utf-16")
			USER_INFO[A_Index, "flags"]            := flags := Format("{:#x}", NumGet(buf + A_PtrSize * 5, "uint"))
			for k, v in UF_FLAGS
				if (k & flags)
					flags_list .= v " | "
			USER_INFO["flags_list"]                := SubStr(flags_list, 1, -3)
			USER_INFO[A_Index, "script_path"]      := StrGet(NumGet(addr + A_PtrSize * 6, "uptr"), "utf-16")
			USER_INFO[A_Index, "auth_flags"]       := auth_flags := Format("{:#x}", NumGet(buf + A_PtrSize * 7, "uint"))
			for k, v in AF_OP_FLAGS
				if (k & auth_flags)
					auth_flags_list .= v " | "
			USER_INFO["auth_flags_list"]           := SubStr(auth_flags_list, 1, -3)
			USER_INFO[A_Index, "full_name"]        := StrGet(NumGet(addr + A_PtrSize * 8, "uptr"), "utf-16")
			USER_INFO[A_Index, "usr_comment"]      := StrGet(NumGet(addr + A_PtrSize * 9, "uptr"), "utf-16")
			USER_INFO[A_Index, "parms"]            := StrGet(NumGet(addr + A_PtrSize * 10, "uptr"), "utf-16")
			USER_INFO[A_Index, "workstations"]     := StrGet(NumGet(addr + A_PtrSize * 11, "uptr"), "utf-16")
			USER_INFO[A_Index, "last_logon"]       := ConvertUnixTime(NumGet(addr + A_PtrSize * 12, "uint"))
			;USER_INFO[A_Index, "last_logoff"]      := NumGet(addr + A_PtrSize * 12 + 4, "uint")   ; currently not used.
			USER_INFO[A_Index, "acct_expires"]     := (NumGet(buf + A_PtrSize * 13, "uint") = TIMEQ_FOREVER)
			                                        ? "NEVER"
			                                        : ConvertUnixTime(NumGet(buf + A_PtrSize * 13, "uint"))
			USER_INFO[A_Index, "max_storage"]      := (NumGet(buf + A_PtrSize * 13 + 4, "uint") = USER_MAXSTORAGE_UNLIMITED)
			                                        ? "UNLIMITED"
			                                        : NumGet(buf + A_PtrSize * 13 + 4, "uint")
			USER_INFO[A_Index, "units_per_week"]   := NumGet(addr + A_PtrSize * 14, "uint")
			USER_INFO[A_Index, "logon_hours"]      := NumGet(addr + A_PtrSize * 15, "uchar*")   ; todo
			USER_INFO[A_Index, "bad_pw_count"]     := NumGet(addr + A_PtrSize * 16, "uint")
			USER_INFO[A_Index, "num_logons"]       := NumGet(addr + A_PtrSize * 16 + 4, "uint")
			USER_INFO[A_Index, "logon_server"]     := StrGet(NumGet(addr + A_PtrSize * 17, "uptr"), "utf-16")
			USER_INFO[A_Index, "country_code"]     := NumGet(addr + A_PtrSize * 18, "uint")
			USER_INFO[A_Index, "code_page"]        := NumGet(addr + A_PtrSize * 18 + 4, "uint")
			USER_INFO[A_Index, "user_id"]          := NumGet(addr + A_PtrSize * 19, "uint")
			USER_INFO[A_Index, "primary_group_id"] := NumGet(addr + A_PtrSize * 19 + 4, "uint")
			USER_INFO[A_Index, "profile"]          := StrGet(NumGet(addr + A_PtrSize * 20, "uptr"), "utf-16")
			USER_INFO[A_Index, "home_dir_drive"]   := StrGet(NumGet(addr + A_PtrSize * 21, "uptr"), "utf-16")
			USER_INFO[A_Index, "password_expired"] := NumGet(addr + A_PtrSize * 22, "uint")
			addr += A_PtrSize * 23
		}

		DllCall("netapi32\NetApiBufferFree", "ptr", buf, "uint")
		return USER_INFO
	}

	DllCall("netapi32\NetApiBufferFree", "ptr", buf, "uint")
	return false
}


ConvertAge(value)
{
	age := A_Now
	age += -(value), s
	FormatTime, age, % age, yyyy-MM-dd HH:mm:ss
	return age
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

; ===============================================================================================================================

NetGroupEnum := NetUserEnum("DC01")
for i, v in NetGroupEnum {
	for k, v in NetGroupEnum[i]
		output .= k ": " v "`n"
	MsgBox % output
	output := ""
}