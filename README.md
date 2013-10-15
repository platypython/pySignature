pySignature
===========
Requires python 2.7, not tested with 3.

Outlook signature generator for Active Directory

Using Tim Golden's excellent active_directory module (http://timgolden.me.uk/python/active_directory.html) pySignature
pulls data from Active Directory and using Jinja2 templates, creates HTML, RTF, and TXT Outlook signatures for 2003, 2007, 2010,
and 2013.

The script is designed to be run as a logon script in the netlogon folder. It will generate the signatures and edit the registry
to automatically enforce them for the current profile. When run as a logon script the signatures will auto-update with any new
information.

Editing the templates is a simple process except for the richtext one. You'll need a little trial and error to get that one perfect.

This differs from the transport rule method because these signatures appear at the end of the most recent post,
instead of at the bottom of the email chain.

I've tested it on XP and 7, hopefully it won't break anything.

Deployment
===========

This can be compiled very easily into an .exe file using the pyinstaller module. Alternatively portable python can be installed in a
network location to run the script.

Requires Tim Golden's active_directory module, Mark Hammond's pywin32 module, and jinja2.