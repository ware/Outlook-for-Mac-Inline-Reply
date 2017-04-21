# Copyright (C) 2017, Ryan Ware
# 
# This program is free software; you can redistribute it and/or
# modify it under the terms of the GNU General Public License
# as published by the Free Software Foundation; either version 2
# of the License, or (at your option) any later version.
# 
# This program is distributed in the hope that it will be useful,
# but WITHOUT ANY WARRANTY; without even the implied warranty of
# MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
# GNU General Public License for more details.
# 
# You should have received a copy of the GNU General Public License
# along with this program; if not, write to the Free Software
# Foundation, Inc., 51 Franklin Street, Fifth Floor, Boston, MA  02110-1301, USA.

tell application "Microsoft Outlook"
	
	set replyMessage to selection
	set MessageSender to sender of replyMessage
	set MessageName to name of MessageSender
	set MessageDate to time sent of replyMessage
	set MessageAddress to address of MessageSender
	set MessageSubject to subject of replyMessage
	set theMaxLineLength to 80
	set theLinePrefix to "> "
	
	if (has html of replyMessage) then
		set theText to do shell script "echo " & quoted form of ("<!DOCTYPE HTML PUBLIC><meta charset=\"UTF-8\">" & content of replyMessage as string) & space & "| textutil -convert txt -stdin -stdout"
	else
		set theText to content of replyMessage as string
	end if
	set theQuotedText to "On " & MessageDate & " " & MessageName & " said:" & return & theLinePrefix
	repeat with a from 1 to (count paragraphs of theText)
		set thePar to paragraph a of theText
		if (length of thePar) is greater than theMaxLineLength then
			set breakLine to false
			repeat with b from 1 to (length of thePar)
				set theChar to character b of thePar
				set theQuotedText to theQuotedText & theChar
				if b mod theMaxLineLength = 0 then set breakLine to true
				if breakLine = true then
					if {tab, space, ".", ",", "?", "!", "\"", "(", ")"} contains theChar then
						set theQuotedText to theQuotedText & return & theLinePrefix
						set breakLine to false
					end if
				end if
			end repeat
		else
			set theQuotedText to theQuotedText & thePar
		end if
		if a is not equal to (count paragraphs of theText) then set theQuotedText to theQuotedText & return & "> "
	end repeat
	
	set newMessage to reply to replyMessage with reply to all
	set the plain text content of newMessage to theQuotedText
	
end tell