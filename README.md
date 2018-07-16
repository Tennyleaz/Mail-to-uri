# Mail-to-uri
A simple utility using mailto: uri protocal. Will also detect if any mail client installed.

## Detecting installed mail clients
`IsMailClientInstalledRoot()` function will look for registry `HKEY_CLASSES_ROOT\mailto\shell\open\command` 
for a command other than default `C:\Windows\System32\run32.dll` or simalar value. 

If a custom value is found, a email client application might installed for all users.

`IsMailClientInstalledForUser()` function does simillar job, the only difference is that it looks for `HKEY_LOCAL_MACHINE\SOFTWARE\Classes\mailto\shell\open\command\`, 
for any email client installed for current user.

## Sending mail

`mailto:` command is used to compose a URI, then let windows system handle it. 

In Windows 8 or later, default mail app is always an option. In Windows 7, if no email client is installed, Internet Explorer will handle this uri.

The address string is encoded by `Uri.EscapeUriString()`, and the title and body strings are encoded by Uri.EscapeDataString().

Note that we want whitespace in email title/body to be encode as `%20`, but not `+`.
