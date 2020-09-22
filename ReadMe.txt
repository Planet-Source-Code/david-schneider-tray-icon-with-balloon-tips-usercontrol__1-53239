Tray Icon UserControl
	by David Schneider

Feel free to use this code, but be warned, it's not commented!
		(It's pretty simple though)

How to use:
	Step 1: Add the included usercontrol to your project.
		It's not an ocx, so you don't have to worry
		about dependencies.

	Step 2: Drag the control from the toolbox onto your form.
		You can add more (or index it) just like a normal
		control.

	Step 3: In the code, when you want your tray icon to appear,
		call the Create command of the user control.

	Step 4: While the tray icon exists, you can call the various
		other functions (such as BalloonTip) or set other 
		properties (such as Icon), as explained	in the 
		reference section.

	Step 5:	Use the supplied events to figure out when the user
		either clicks on the tray icon or the balloon--
		such as creating a popup menu on right mouse button up
		or making your form appear on double-click

	Step 6: Call the Remove function to remove the tray icon.

Reference:
	Really, the stuff is self-explanitory, and I made sure to not
	overwhelm the user, but whatever, here's the info about the
	usercontrol's properties, functions, and events.

Functions:
	.Create ([ToolTipText][, Icon as StdPicture])

		Creates the system tray icon.
		ToolTipText sets the caption that pops up when you
			hold your mouse over the icon.
		Icon sets the icon of the system tray icon.
			Not supplied = transparent

	.Remove

		Destroys the tray icon.  You can always Create it again.

	.BalloonTip (Prompt As String[, Style As BalloonTipStyle] _
		          [, Title As String][, BTimeout As Long])
		
		Creates a non-obtrusive balloon tip popup, provided that the
			OS supports it and balloon tips are not disabled.
		Prompt is the smaller message text of the balloon tip.  Can
			be multi-line.
		Style sets the icon in the top-left corner of the balloon tip.
			They are the miniature counterparts to the MessageBox's.
			If not supplied then the icon is blank.
		Title sets the bold caption at the top of the balloon tip.
			If not supplied or empty then the caption is the
			application's title.
		Timeout sets the timeout for the balloon tip.  Personally, I have
			not found this to do anything on XP at least, so
			generally you can just leave it at the default, and it's
			there because it's an important part of the API call.
	
	.PopupMenu (Menu As Menu, Optional Flags, Optional DefaultMenu)
		
		Creates a popup menu where the mouse is.  Generally you use this
			in the TrayClick event.  Use this INSTEAD of the normal
			VB PopupMenu function because this one makes the menu
			dissapear if you click off of it.
		Menu is the menu object that you want to pop up.  It must have
			at least one submenu.
		Flags are the flags for the popup menu.  These are the same as
			the standard popup menu's flags, as defined in the
			enumeration MenuControlConstants.  Generally you can
			leave this empty.
		Default menu is the menu that is in bold when the menu pops up.
			This MUST be a submenu of Menu, and MUST be visible.
			From a GUI perspective, this menu is generally the same
			one that is clicked when you just double-click the tray
			icon.
	
Events:
	TrayClick(Button As stMouseEvent)

		Is called when the user clicks the tray icon in any way.
		Button is the type of click that occurred.  This can be the
		left or right mouse button, and a mouse down, mouse up, or 
		double-click.  Note that the standard click variables ARE THE SAME
		as the mouse up.
	
	BalloonClick(ClickType As stBalloonClickType)
	
		Is called when the user clicks on a popped-up balloon.
		ClickType is the type of click--general left-click, right-click, or
		left-click on the X button (XP).  Note that right-click and X-click
		ARE THE SAME.

Properties:
	Icon
		The icon of the tray icon.  Note that setting this causes the balloon
		tip to dissapear if it's up at the time.

	ToolTipText
		The text that appears when you hover over the tray icon.  Note that
		setting this causes the balloon tip to dissapear if it's up at the time.


That's all!  Hope this really improves your programs!

E-mail me at Davidnsc1@aol.com with comments, or just post something at PSC.
PLEASE e-mail me if you didn't find this at Planet Source Code, as this should ONLY be
available at that great website that has helped me many times in the past.

Don't forget to rate it if you liked it!  (I guess you can rate it if you didn't, but...)