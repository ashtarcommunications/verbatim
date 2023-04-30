---
sidebar_position: 2
id: virus
title: Verbatim and Antivirus
---

# Verbatim and Antivirus

Verbatim is often flagged as a false postive by antivirus programs, including Microsoft Defender. This is because it makes use of Visual Basic for Applications (VBA) Macros, which are also commonly used by some kinds of viruses and trojans. Verbatim also includes code for things like sending files to the caselist or checking for updates, which look suspicious to the heuristics used by automated scanners to detect malware.

There is nothing harmful in Verbatim, and it's not a virus. It's completely open source, so you're welcome to inspect the code yourself to verify.

Each new release of Verbatim is submitted to antivirus companies for approval, so the false postives are usually temporary. But if you're still getting a false positive, your best bet is to tell your antivirus software to allow Verbatim. You may need to reinstall Verbatim after doing so.

## Windows Defender

First, ensure you've updated Microsoft Defender to the latest security definitions. Verbatim is usually marked as allowed in the most recent definitions.

However, Microsoft Defender has recently become significantly more aggressive at flagging files with VBA macros, even when harmless. It may even remove the macros from the Debate.dotm template and block Verbatim without notifying you. After installing Verbatim, check your Windows Defender protection history:

![Windows Defender](../assets/windows-defender.png)

If you have an alert about Debate.dotm, use the Actions menu for the alert to "Allow" or "Restore" the file, depending on what the scanner did to block it.
