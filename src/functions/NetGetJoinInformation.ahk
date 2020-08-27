; ===============================================================================================================================
; Function ...: NetGetJoinInformation
; Return .....: Retrieves join status information for the specified computer.
; Link .......: https://docs.microsoft.com/en-us/windows/win32/api/lmjoin/nf-lmjoin-netgetjoininformation
; ===============================================================================================================================

NetGetJoinInformation(Server := "127.0.0.1")
{
	static NERR_SUCCESS := 0
	static JOIN_STATUS := { 0: "Unknown", 1: "Unjoined", 2: "Workgroup", 3: "Domain" }

	NET_API_STATUS := DllCall("netapi32\NetGetJoinInformation", "wstr", Server
	                                                          , "ptr*", BufferName
	                                                          , "int*", BufferType
	                                                          , "uint")

	if (NET_API_STATUS = NERR_SUCCESS)
	{
		JOIN_INFO := []
		JOIN_INFO["Name"] := StrGet(BufferName)
		JOIN_INFO["Type"] := JOIN_STATUS[BufferType]

		DllCall("netapi32\NetApiBufferFree", "ptr", buf, "uint")
		return JOIN_INFO
	}

	DllCall("netapi32\NetApiBufferFree", "ptr", buf, "uint")
	return false
}

; ===============================================================================================================================

for k, v in NetGetJoinInformation()
	output .= k ": " v "`n"
MsgBox % output