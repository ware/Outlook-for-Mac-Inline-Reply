# Outlook-for-Mac-Inline-Reply
This small project contains an AppleScript that people can use to do quoted inline plaintext replies in Outlook for Mac.  Unfortunately, it hasn't been possible to do this in Outlook for Mac since they switched their editing back end to Word at version 15.22.  This AppleScript give a way to do this.

Note that this is currently nothing but a hack by someone who has avoided AppleScript like the plague all their life.  I've culled this together from various examples found in various web forums demonstrating key pieces.  Don't expect it to be anything other than a hack.  Non-hacky pull-requests happily reviewed!

# "Install"
There is no real set install.  I can convey how I utilize this.

First, you will want to open "Reply Plain Text.applescript" in Script Editor.  Once opened in there, you will want to compile it (hammer icon), then "File"->"Export..." and save it as a script.

I save the resultant script file in ~/Library/Scripts.

There is unfortunately no way to execute this directly from within Outlook for Mac itself.  To execute this with a key combination, I use [FastScripts](https://red-sweater.com/fastscripts/) from [Red Sweater Software](https://red-sweater.com).  It allows you to assing shortcuts to specific AppleScripts.

There is no doubt that there are numerous ways to go about doing this.  This is just one.