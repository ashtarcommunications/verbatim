---
sidebar_position: 5
id: faq
title: FAQ
---

# Frequently Asked Questions

## Installation

### Why do I get a virus alert when I download Verbatim?
See the section on [Virus Scanners](./getting-started/virus). You can also try Verbatim Mini, which is a stripped-down version of Verbatim with less features likely to trip your virus scanner. For more information, see [Installation](./getting-started/installation).

### I'm getting a message that Debate.dotm doesn't exist after installation.
You’re most likely using the version of Office in the Microsoft App Store, which doesn’t fully support VBA Macros. Try installing a full version of Office. For more details, see [Requirements](./getting-started/requirements).

### I can't run the Verbatim Installer on Mac.
Usually, this is because your security settings prohibit running installers from the internet. To fix, try running the Verbatim installer package, then go to the Apple menu – System Preferences – Security & Privacy. You’ll see a notice asking whether to allow installation to proceed.

### What versions of Word work with Verbatim?
See [Requirements](./getting-started/requirements).

### Will Verbatim work on my tablet?
Only if your tablet runs a full version of Microsoft Office with support for VBA macros.
That means it won’t work on an iPad or an Android tablet, but should work on a Surface Pro. It also won’t work on a Surface RT – the version of Office that runs on the RT doesn’t include VBA macro support.

### Will Verbatim work on a Chromebook?
No. Verbatim requires a full version of Word with support for VBA macros, which currently doesn’t run on a Chromebook.

### Will Verbatim work in OpenOffice or Pages?
No, Verbatim is built as a template for Microsoft Word, so requires a full version of Word to function.

### How do I completely uninstall Verbatim?
On PC, just run the uninstaller from Add/Remove programs. On Mac, you can use the UninstallVerbatim.zip app package from the Downlaods page.

Alternately, to perform a manual uninstall, delete the files you installed as part of a [Manual Installation](./getting-started/installation).

### I ran the automatic installer – how do I know if it’s installed right?
You should have a shortcut to Verbatim on your desktop, or on Mac, in your Applications Folder. If this doesn’t appear, you can also check to see if it’s installed by opening Word and selecting File – New, and looking for a Debate template in your Templates section.


## Security

### How do I make security prompts and protected view warnings go away?
See the section on [Security](./getting-started/security)


## Verbatim Usage

### My macros aren’t working at all
Most of the time, this is because your macro security settings are set too high, or your antivirus software has blocked Verbatim. See the sections on [Security](./getting-started/security) and [Antivirus](./getting-started/virus).

You should also check to make sure that you have Verbatim installed correctly in the Word Templates folder – otherwise someone else’s Verbatim document won’t be able to find it.

### I'm getting permissions errors or "Error 5487 - can't save changes to Debate.dotm" when trying to save settings
This is due to Verbatim being blocked by Windows Defender. See the section on [Antivirus](./getting-started/virus).

### I get "Error 5152: Invalid Document Name" when trying to save a speech document
Most likely, this is due to having your Auto Save folder set to a cloud folder (e.g. Dropbox, Google Drive, OneDrive), which isn't compatible with auto saving. Try changing your Auto Save folder to a local folder, like your Desktop or Documents folder.

### I get "Error 5852: Requested Object Is Not Available" when working with documents
This is simlar to the 5152 error described above, due to working with documents stored in a cloud folder (e.g. Dropbox, Google Drive, OneDrive). Try moving the document to a local folder and seeing if the error goes away.

### I get "Error 52: Bad File Name or Number" while uploading open source documents to the caselist
This can happen when the file you are uploading is saved in a folder synced with an online cloud storage service like OneDrive or Dropbox. It happens because Word is unable to access the file from a networked filesystem while converting it into a format that the caselist can read. As a workaround, try temporarily saving a copy of your file to a non-synced folder, and uploading that copy instead.

### My document keeps jumping to the top of the page on Mac
This is due to a bug in Mac Word, not in Verbatim. It occurs if you’re in Web view, with the Navigation Pane turned on. The easiest fix is to switch to Draft view, or turn off the Navigation Pane. You can set your default view to Draft in the Verbatim settings.

### Why don't the F-key shortcuts work on my Mac?
First, ensure that you have "Use Fn keys as standard function keys" checked in your System Preferences - Keyboard settings.

If some F-keys still don't work, it's likely you have conflicting Mac OS keyboard shortcuts preempting the Verbatim keys. You can disable these in your System Preferences, or run the Verbatim Setup Tool, which can automatically configure them for you.

Note that the F6 key may be impossible to get to work - on some versions of Mac Word, F6 appears to be hardwired to a "Switch Pane" command, and isn't changeable by Verbatim. In that case, it's suggested you use the alternate keyboard shortcut, `Cmd + Alt + 6`

### My Debate ribbon is collapsing into sections
Either you have a small screen, or you have “Touch Mode” turned on in Word. Turn it off by following the directions here:

https://support.office.com/en-us/article/turn-touch-mode-on-or-off-90e162b1-44f9-434f-bc2b-9321c989ea6e

### Why does Gmail mark my Verbatim files as having a virus?
Gmail sometimes marks files containing VBA macros as a security risk. There is unfortunately nothing you can do – Gmail is in complete control of what they mark as a virus and what they don’t. If the file is a docx, then it doesn't contain any macros or code (all code lives inside the master template on your computer), so it’s impossible for them to have a virus.

Your best bet is to use something other than email to transfer your files, or to copy and paste the contents of your verbatim file into a new non-Verbatim blank Word document before sending via email.

### The Setup Wizard opens every time I open Verbatim on Mac
If the Setup Wizard opens every time you open Word, it’s likely due to a permissions issue with your Verbatim preferences file, usually from a previous installation. To fix, just delete your Verbatim.plist file, which is usually located at `~/Library/Preferences/Verbatim.plist`

If your Library folder is hidden by default, you can press `Cmd + Shift + G` to open the “Go” menu, and type `~/Library/Preferences`

Then, delete Verbatim.plist

### Word keeps prompting me to save changes to Debate.dotm, even though I haven’t actually changed anything.
This problem is most likely unrelated to Verbatim, and is caused by an error in a 3rd party “COM Add-In” that you probably installed inadvertently. Usually, this is caused by one called “Send to Bluetooth.” The solution is to disable all Word Add-In’s, and see if the behavior goes away.

To disable, go to File – Word Options – Add-Ins. In the “Manage” drop-down box, select “COM Add-Ins.” You should then uncheck any Add-ins, especially if you have one called “Send to Bluetooth,” and click OK.

Alternately, try running the Verbatim Troubleshooter and see if it lists any installation problems.

### I get an error "User-defined type not defined for Private Function FindVBProject" on the Mac
This is caused by having older incompatible Mac Verbatim versions still installed. Try opening your Templates folder and deleting your Normal.dotm template, then reinstalling Verbatim.

### I pressed the macro hotkey, and my screen suddenly rotated 90 degrees.
This occurs on certain laptops using a particular graphics card software package. To get your screen back to normal, press Ctrl-Alt-Up Then, right click on your desktop, select “Graphics Options – Hotkeys” and select “Disable hotkeys.”

### Why is there no "Card" style?
Because there’s not much in the document which isn’t already a different heading which is NOT card text. This is done to simplify many of the macros, and because it makes sense to identify the contents of a card by the associated Tag. Since every card comes with a Tag, a separate style for Card is unnecessary.

If what you’re really looking to do is just indent you card text, that function is already built into Word – by default the keyboard shortcut is Ctrl+M.

### Why is there no "Analytics" style?
The reason that people want an Analytics style is almost exclusively so they can write a macro to delete the analytics before sharing with their opponent or judge. While I understand this impulse, I think that it's misguided and should be discouraged.

Any time you make the arguments made in your speech less accessible to your opponent, you're running away from clash and encouraging worse debate. Removing content from your speech document is anti-competitive, anti-educational, and damaging for people who may need the written document to engage with your arguments.

For similar reasons, Verbatim will never support shrinking text smaller than 4pt (which is already probably too small), or versions of Invisibility Mode which actually delete text from the document.

Debaters should win with superior argumentation, not cheap tricks.

### Every time I paste text, it comes in as totally unformatted – even when I just use Ctrl+V.
This is not a Verbatim-specific issue, but it is one of the strangest behaviors I’ve ever seen in Word. The solution is to uninstall a program called “Skype Click-To-Call.” You may not even have installed this on purpose – it is sometimes included as part of an automatic Skype upgrade. Go to Control Panel – Add/Remove Programs and uninstall it – your cut and paste functionality should return to normal immediately. Don’t ask me to even kinda explain this one.


## In-Round

### What do I do if my computer crashes?
First, minimize the chance of a crash by practicing good preventative care on your computer.  Ensure your operating system is up to date, take care of it physically, etc. That said, if it does happen, there are several backups in place. Since each debater should have a copy of the speech there should be several other computers looking at the current speech at any given time. If a computer crashes before the speech, a reboot will usually solve the problem – and if the debater has been saving regularly, not much work should be lost.

### How do I mark a card while giving a speech?
While in your Speech doc, click on the part of the card where you stopped and press the `\~ key, which will insert a marker. Most debaters have also taken to saying “marked at xxxxxx” when marking to let the other team know where they stopped if they’re following along.


## Customization

### How do I change the macro hotkeys?  
You can set Word to use any key combination you choose for each macro in lieu of the default hotkeys. The preferred way to do this is to use the built-in customization interface in the Verbatim settings. For more extensive changes, you can use the "Other Shortcuts" button to open Word's keyboard customization interface. Make sure you have "Save Changes in..." set to Debate.dotm if you want your changes to be persistent. Keep in mind that the `/~ key cannot be assigned using this method.
