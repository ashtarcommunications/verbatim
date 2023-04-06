# Frequently Asked Questions

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
