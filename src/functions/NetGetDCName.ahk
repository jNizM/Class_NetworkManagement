; ===============================================================================================================================
; Function ...: NetGetDCName
; Return .....: Tthe name of the primary domain controller (PDC). It does not return the name of the backup domain controller (BDC)
;               for the specified domain. Also, you cannot remote this function to a non-PDC server.
; Link .......: https://docs.microsoft.com/en-us/windows/win32/api/lmaccess/nf-lmaccess-netgetdcname
; ===============================================================================================================================

NetGetDCName(ServerName := "", DomainName := "")
{
	static NERR_SUCCESS := 0

	NET_API_STATUS := DllCall("netapi32\NetGetDCName", "wstr", ServerName
	                                                 , "wstr", DomainName
	                                                 , "ptr*", buf
	                                                 , "uint")

	if (NET_API_STATUS = NERR_SUCCESS)
	{
		DomainController := buf
		DllCall("netapi32\NetApiBufferFree", "ptr", buf, "uint")
		return StrGet(DomainController)
	}
	else
	{
		DllCall("netapi32\NetApiBufferFree", "ptr", buf, "uint")
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

; ===============================================================================================================================

MsgBox % NetGetDCName()