# Class_NetworkManagement
 AutoHotkey wrapper for Network Management technology ([msdn-docs](https://docs.microsoft.com/en-us/windows/win32/api/_netmgmt/))


## Info
This wrapper could be useful for administrators and first and second level supporters.
All functions in this class are also available as standalone functions. ([src/functions](src/functions))


## Examples

**Retrieves a list of global groups to which a specified user belongs.**
```AutoHotkey
for k, v in NetworkManagement.NetUserGetGroups(A_UserName, "DC01")
	output .= k ": " v "`n"
MsgBox % output
```

**Retrieves a list of local groups to which a specified user belongs.**
```AutoHotkey
for k, v in NetworkManagement.NetUserGetLocalGroups(A_UserName, "DOMAIN\", "DC01")
	output .= k ": " v "`n"
MsgBox % output
```

**Retrieves join status information for the specified computer.**
```AutoHotkey
for k, v in NetworkManagement.NetGetJoinInformation()
	output .= k ": " v "`n"
MsgBox % output
```

**Returns information about the configuration of a workstation**
```AutoHotkey
for k, v in NetworkManagement.NetWkstaGetInfo()
	output .= k ": " v "`n"
MsgBox % output
```

**Lists information about all users currently logged on to the workstation. This list includes interactive, service and batch logons.**
```AutoHotkey
NetWkstaUserEnum := NetworkManagement.NetWkstaUserEnum()
for i, v in NetWkstaUserEnum {
	for k, v in NetWkstaUserEnum[i]
		output .= k ": " v "`n"
	MsgBox % output
	output := ""
}
```

**Returns information about the currently logged-on user.**
```AutoHotkey
for k, v in NetworkManagement.NetWkstaUserGetInfo()
	output .= k ": " v "`n"
MsgBox % output
```


## Questions / Bugs / Issues
Report any bugs or issues on the [AHK Thread](https://www.autohotkey.com/boards/viewtopic.php?f=6&t=80382). Same for any questions.


## Copyright and License
[The Unlicense](LICENSE)


## Donations (PayPal)
[Donations are appreciated if I could help you](https://www.paypal.me/smithz)