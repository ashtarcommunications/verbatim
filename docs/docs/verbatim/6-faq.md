---
sidebar_position: 6
id: faq
title: FAQ
---

# Frequently Asked Questions


I’ve heard the built-in timer has a bitcoin miner?

No timer is included in the most recent versions of Verbatim (v5.1.1). A brand new cross-platform open source timer will be released alongside Verbatim 6.

In earlier versions of the installer package (PC only), Verbatim bundled the “Debate Synergy” timer v. 1.5 created by Alex Gulakov as part of the Debate Synergy template and released for free under the GPL 3. That software was released in 2010 and the version that was included in Verbatim has not been modified since that time. I had nothing to do with its creation.

There was at one point an unsubstantiated rumor that the Synergy timer contained a bitcoin miner. I have no evidence that this rumor has any basis in fact. Alex Gulakov has also contacted me to say that this claim is false. He provided a link to the following code repository:
https://github.com/vtempest/DebateSidebar/tree/master/DebateTimer

If you have an older PC version and remain concerned about this, you can uninstall the Synergy timer by deleting Timer.exe in your Templates folder.

I'm getting a message that Debate.dotm doesn't exist after installation.

You’re most likely using the version of Office in the Microsoft App Store, which doesn’t fully support VBA Macros. Try installing a full version of Office.

Mac – My document keeps jumping to the top of the page

This is due to a bug in Mac Word, not in Verbatim. It occurs if you’re in Web view, with the Navigation Pane turned on. The easiest fix is to switch to Draft view, or turn off the Navigation Pane. You can set your default view to Draft in the Verbatim settings.

Mac – The Verbatim/Add-Ins toolbar doesn't open with older documents.
This is a problem/limitation with how Mac Word handles global templates, not a problem with Verbatim.
The bottom line is there is very little you can do – if you want Verbatim to work on all your documents, consider using Boot Camp/Paralells. PC Word doesn’t have this problem.
There are some things you can try, but your mileage may vary. Mac Word’s support for global templates is very spotty, and it’s unlikely to work perfectly with files produced on other computers, especially files
originally produced on a PC. The usual workaround is to open a new Verbatim doc, then cut and paste the contents.
Other things you can try:
1) Make sure the Verbatim.scpt file is installed in this directory:
~/Library/Application Scripts/com.microsoft.Word/
2) Try deleting your Verbatim.plist file in ~/Library/Preferences, and
your Normal.dotm template in:
~/Library/Group Containers/UBF8T346G9.Office/User Content/Templates
Then, repeat the “Verbatimize Normal” steps in the Verbatim Settings.
If your hard drive isn’t named the standard “Macintosh HD” you can also
try changing it back to the default and see if that helps.

My Debate ribbon is collapsing into sections
You have “Touch Mode” turned on in Word. Turn it off by following the directions here:
https://support.office.com/en-us/article/turn-touch-mode-on-or-off-90e162b1-44f9-434f-bc2b-9321c989ea6e

What versions of Word work with Verbatim?
Windows – Verbatim 5.1+ will work with Word 2010, Word 2013, or Word 2016.
Mac – Verbatim 5.2+ will work with Mac Word 2011 or Mac Word 2016.
Verbatim requires a “full” version of Microsoft Office (Office 365 or regular), which includes support for VBA macros. It will not work with the Office Starter Pack, Office Home and Student, Office RT (e.g. for the Surface tablet), Office for iPad, Office for Android, or the version of Office in the Microsoft App Store.

Why does Gmail mark my Mac Verbatim files as having a virus?
Gmail recently changed their virus scanner, and are now reporting false positives for files produced in Mac Verbatim. There is unfortunately nothing I can do – Gmail is in complete control of what they mark as a virus and what they don’t. The files being sent do not contain any macros or code (all code lives inside the master template on your computer), so it’s impossible for them to have a virus. That is, there’s no way to modify Mac Verbatim to avoid this issue.
Your best bet is to use something other than email to transfer your files, or to copy and paste the contents of your verbatim file into a new non-Verbatim blank Word document before sending via email.

Will Verbatim work on my tablet?
Only if your tablet runs a full version of Microsoft Office with support for VBA macros.
That means it won’t work on an iPad or an Android tablet, but should work on a Surface Pro. It also won’t work on a Surface RT – the version of Office that runs on the RT doesn’t include VBA macro support.

Will Verbatim work on a Chromebook?
No. Verbatim requires a full version of Word with support for VBA macros, which currently doesn’t run on a Chromebook.

Will Verbatim work in OpenOffice or Pages?
No, Verbatim is built as a template for Microsoft Word, so requires a full version of Word to function.

How do I completely uninstall Verbatim?

Windows:
Just run the uninstaller from Add/Remove programs. If you were using Verbatim’s “Always On” mode, you can use the “Unverbatimize Normal Template” button in the Verbatim Settings before running the uninstaller.
Alternately, to perform a manual uninstall on Windows you just need to delete the Debate.dotm file and your Normal.dotm file from your Templates folder. Word will automatically regenerate a new Normal.dotm file. The Templates folder is usually located at:
c:\Users\[Your Name]\AppData\Roaming\Microsoft\Templates
Mac:
On the Mac, you can uninstall by deleting Debate.dotm and Normal.dotm from your My Templates folder.
If you’re using Mac Word 2011, the files are located at:
~/Library/Application Support/Microsoft/Office/User Templates/
and in the My Templates subfolder.
If you’re using Mac Word 2016, they’re in:
~/Library/Group Containers/UBF8T346G9.Office/User Content/Templates
You can also delete the Verbatim.scpt file (if present) in:
~/Library/Application Scripts/com.microsoft.Word/
You can use Cmd+Shift+G to go directly to the folders above.

Mac – I can't run the Verbatim Installer.
Usually, this is because your security settings prohibit running installers from the internet. To fix, try running the Verbatim installer package, then go to the Apple menu – System Preferences – Security & Privacy. You’ll see a notice asking whether to allow installation to proceed.
Alternately, you can follow these instructions to change your system settings to allow installation.

Mac – The Setup Wizard opens every time I open Verbatim.
If the Setup Wizard opens every time you open Word, it’s likely due to a permissions issue with your Verbatim preferences file, usually from a previous installation.
To fix, just delete your Verbatim.plist file, which is usually located at ~/Library/Preferences/Verbatim.plist
If your Library folder is hidden by default, you can press:
Cmd + Shift + G to open the “Go” menu, and paste:
~/Library/Preferences
Then, delete Verbatim.plist

Why do I get a virus alert when I download Verbatim?
Some virus scanners (especially Microsoft and Norton/Symantec) may give a false positive virus alert when downloading or trying to use Verbatim. Recently, Microsoft in particular has gotten much more aggressive at flagging false positives with files that contain VBA macros.

Don’t worry, it’s completely safe – there is nothing harmful in Verbatim. False positives are just due to auto-detection rules when an installer contains a file with macros. If you’re still concerned, Verbatim is open source – you can read every line of code in it if you’re so inclined or perform a manual installation rather than using the installer package.

Each new release of Verbatim is submitted to antivirus companies for approval, so the alerts usually go away.
In the interim, you’re best off using your antivirus settings to “whitelist” Verbatim. You may need to reinstall Verbatim after doing so.


## Paperless Concerns

### What do I do if my computer crashes?
First, keep in mind that it’s important to minimize the chance of a crash by practicing good preventative care on your computer.  Ensuring your operating system is up to date, that you’re running anti-virus software, and that the machine is physically well taken care of will go a long ways towards avoiding any problems. 
That said, if it does happen, there are several backups in place.  Since each debater should have a copy of the speech there should be several other computers looking at the current speech at any given time. If a computer crashes before the speech, a reboot will usually solve the problem – and if the debater has been saving regularly, not much work should be lost.

### Can this work on Linux, or with Open Office?
Not yet.  Open Office isn’t good enough yet.  It’s pretty close to replicating most of the needed functionality, but support for macros is still pretty lagging, and it’s lacking “Draft View,” which is a deal breaker. Hopefully in years to come this becomes a more viable option. 

I’m also very confident that porting the whole VBA code base into Open Office’s native macro format shouldn’t be that hard.  In many ways, it’s better than VBA – but the other failings of Open Office make this a very low priority right now.

## Installation Problems

### I ran the automatic installer – how do I know if it’s installed right?
On the PC, you should have a shortcut to Verbatim on your desktop. If this doesn’t appear, you can also check to see if it’s installed by opening Word and selecting File – New, and looking for Debate.dotm in your “My Templates” section.

On the Mac, there will be no immediately apparent sign that Verbatim installed correctly. The only way to check is to start Word, go to File – New from Template, and see if Debate appears under My Templates.

### Why can’t I just install Verbatim as my Normal template or leave it on my desktop?
Because for files you produce to be recognized as Verbatim documents by other people, you both need to have the same template, in the same place, with the same. When you install Verbatim as your “Normal” template, you are breaking compatibility with everyone else.

Verbatim is also not designed to be your “default” template when you open Word. That’s because most people that use Word still use it more for non-debate stuff (like paper writing or business stuff, etc.) and don’t always want to open every document with that template.

The solution to wanting Verbatim to be “always available” is to make use of the Verbatimize button, as described elsewhere in this manual.

For similar reasons, it is highly recommended that you put the template in your Templates folder, not just on your desktop or in your Startup folder.  This is because all documents based on a template don’t actually include the macros from the original.  Instead, they include a “reference” to what template they are based on.  So if you start a new blank document from a template, it will think it is tied to the file on your desktop, not to the one in your Templates folder.  If you then send your file to someone else on your team, it will fail to find the template on their desktop, and the macros will be “missing.”  The obvious solution is for everyone to keep the template installed in the same location.  After installing, just put a shortcut to the file in your Templates folder on your desktop.  Upgrading or changing the template is then as simple as replacing the old version with the updated one in the Word template folder.

If you want a shortcut on your desktop, you should create a shortcut to Debate.dotm that opens a new blank document based on that template, and use that to start a new “Verbatim document” each time, instead of the normal Word icon. If you right-click and drag Debate.dotm to the desktop from your Templates folder you will get an option to “create shortcut here.” If you’re running Windows 7 you can also drag Debate.dotm to your Word icon on the taskbar and “pin” the file as a shortcut for easy access. Running the installation package for Verbatim 4 should have also created that shortcut for you. You can also manually open a new Verbatim doc within word by using the File – New menu.

### How do I uninstall the Verbatimize button?
To completely uninstall the Verbatimize button, delete DebateStartup.dotm in your Word STARTUP folder.


### Word keeps prompting me to save changes to Debate.dotm, even though I haven’t actually changed anything.
This problem is most likely unrelated to Verbatim, and is caused by an error in a 3rd party “COM Add-In” that you probably installed inadvertently. Usually, this is caused by one called “Send to Bluetooth.” The solution is to disable all Word Add-In’s, and see if the behavior goes away.

To disable, go to File – Word Options – Add-Ins. In the “Manage” drop-down box, select “COM Add-Ins.”

You should then uncheck any Add-ins, especially if you have one called “Send to Bluetooth,” and click OK.

You should then repeat these steps, selecting “Word Add-Ins” instead of COM Add-Ins.

## General Macro Problems

### My macros aren’t working at all.
Most of the time, this is because your macro security settings are set too high.  See the Installation section for more specific information on how to enable macros.

You should also check to make sure that you have Verbatim installed correctly in the Word Templates folder – otherwise someone else’s Verbatim document won’t be able to find it.

### My macros keep “disappearing” from my document.
There are a variety of reasons why you might open a document originally created in Verbatim and find the macros and Debate tab missing.  Most commonly, this is either because:
a) You have Verbatim installed incorrectly
or
b)The file was produced on a computer where Verbatim was installed incorrectly

Make sure that you double-check the correct installation steps, and that the template is both named Debate.dotm and located in the correct Word Templates folder. More complete instructions can be found in Chapter 1.

The quickest remedy to an individual file that appears to be missing Verbatim is to use the “Verbatimize” button (PC) or the “Attach Verbatim” toolbar (Mac) to quickly turn the document into a Verbatim file.

Other possible culprits include:
•	Macro security settings – make sure these are turned to low or "off" in Word.
•	Sending a file through email – gmail and other software can strip macros for security reasons.
•	Saving as the wrong type of file – you should always save as .docx files
•	Some other security program like anti-virus, anti-spyware, etc...

### I emailed a file to another team member, and the macros stopped working.
Some email programs or online mail services have been found to strip all macros from Word files when sending them as an attachment, presumably as a security “feature.”  If you find this happening to you, try sending the Word document in a zip file, or with a temporarily modified file extension, such as File.dco instead of File.docx.  

### I pressed the macro hotkey, and my screen suddenly rotated 90 degrees.
This occurs on certain laptops using a particular graphics card software package.  To get your screen back to normal, press Ctrl-Alt-↑.  Then, right click on your desktop, select “Graphics Options – Hotkeys” and select “Disable hotkeys.”

### How do I change the macro hotkeys?  
You can set Word to use any key combination you choose for each macro in lieu of the default hotkeys.  The preferred way to do this is to use the built-in customization interface in the Verbatim settings. This is described in details elsewhere in the manual.

To make other changes, you should open the actual template file (Debate.dotm) and then go to Word Options – Customize the Ribbon and then press the button for “Customize Keyboard.”  Ensure Debate.dotm is selected in the “Save Changes In” box, and then scroll down in the left box to “Styles” and “Macros” – the macros and styles that appear in the right box will then list their associated keyboard shortcuts and allow you to change them.  Keep in mind that the `/~ key cannot be assigned using this method.  Neither can Ctrl-Tab.

### Word stops responding with one of the macros – all I get is an hourglass.
This is probably caused because a macro is in an infinite loop.  I’m pretty sure this won’t happen, because any circumstance where an infinite loop is possible has been coded around.  But, if all else fails, you can manually stop a macro by pressing Ctrl-Break.  Break is also sometimes labeled “Pause” on the keyboard.

### Every time I open a document, I get a prompt that the file is in “Protected View,” which prevents me from editing it.
This is an annoying “security” measure in Word that it’s relatively easy to turn off. Go to File – Word Options – Trust Center – Trust Center Settings, and then Protected View. Uncheck all of the available options and click OK.
 
## Specific Macro Errors

### When I type my apostrophe key, it tries to Send To Speech – my keyboard confuses the ‘ with the `
This bug occurs very rarely, and seems mostly limited to Macbooks. Recent code changes have limited the scope of this bug even further, but if it’s still happening to you, your best bet is to manually change the keyboard shortcut for the Send To Speech macro. You’ll have to edit the AutoOpen and AutoNew macros to remove the lines that start with “KeyBindings.Add” and then manually assign a different keyboard shortcut in the Customize Keyboard interface. Additional fixes to this will happen in the next version of Verbatim.

### Why is there no “Card” style?
Because there’s not much in the document which isn’t already a different heading which is NOT card text. This is done to simplify many of the macros, and because it makes sense to identify the contents of a card by the associated Tag. Since every card (pretty much) already comes with a Tag, a separate style for Card is inefficient.

If what you’re really looking to do is just indent you card text, that function is already built into Word – by default the keyboard shortcut is Ctrl+M.

### Every time I paste text, it comes in as totally unformatted – even when I just use Ctrl+V.
This is not a Verbatim-specific issue, but it is one of the strangest behaviors I’ve ever seen in Word. The solution is to uninstall a program called “Skype Click-To-Call.” You may not even have installed this on purpose – it is sometimes included as part of an automatic Skype upgrade. Go to Control Panel – Add/Remove Programs and uninstall it – your cut and paste functionality should return to normal immediately. Don’t ask me to even kinda explain this one.

### When I use the Condense macro on text from a PDF with columns, it puts in too many Pilcrow signs.
This is a known limitation on the Condense macro – a workaround will be included in a future version. For now, the best option is to open the Verbatim settings and temporarily disable the “Retain Paragraph Integrity” option until you’re done cutting that article.

### I re-opened my file after saving it, and all my highlighting was gone.
This isn’t caused by Verbatim, but is a known bug in Word for which there is no permanent solution.
The first thing to try is highlighting in a different color. Light gray (25%) is more prone to this bug than other colors - so if you're using grey, try changing your default highlighting color and see if the problem disappears.

## In-Round

### How do I mark a card while giving a speech?
While in your Speech doc, click on the part of the card where you stopped and press the `\~ key.  It will insert a marker, like this:
 
Most debaters have also taken to saying “marked at xxxxxx” when marking to let the other team know where they stopped if they’re following along.

### I sent something to my Speech document, and now it looks weird (bigger font, out of order, etc...)
Keep in mind that the “Send To Speech” macro will send your selection to the current insert point in the Speech document.  Odds are good that you accidentally had the cursor in a Block Title or other area of formatted text – and Word attempted to apply that formatting to everything you sent.  Try pressing Undo and then resending the blocks to the bottom.  This is now largely error-trapped to prevent you from doing it – but it’s still possible to get around.

### Can I flow on my laptop?
In short, yes.  With a little practice you should be able to use Alt-Tab to switch between your flow and your Speech document during the speech.  Another suggestion that I haven’t tried personally is to keep two columns of your flow open on the far left of the screen, and put the Speech document to the right so you can see both at the same time.

Nonetheless, I recommend to my students that they flow on paper.  I think that it helps them to see connections between arguments and focus on the big picture.  It also helps to minimize computer distractions during the speech.
