# TheIPChanger

*"New Age Ip Changer", version 0.1.0 (BETA)*

One of the first programs I ever worked on: a Tibia IP changer in Visual Basic 6, which I built on and extended on a Windows XP machine in 2006–2007, when I was 11–12 years old. I didn't write all of it; see [Who wrote what](#who-wrote-what).

## What it does

Tibia's client was hard-wired to CipSoft's login servers. To play on an [OpenTibia](https://en.wikipedia.org/wiki/OpenTibia) server (a fan-made server emulator), you needed an "IP changer": a small tool that finds the running Tibia window, opens its process, and uses `WriteProcessMemory` to overwrite the login-server hostname in the client's memory with the server you want.

TheIPChanger does that for clients **7.1, 7.24, 7.3, 7.5, 7.6 and 7.91**, with the memory address for each version hard-coded (7.1 needed a second address for the port). It also has:

- a saved server list with add/delete/export
- a "Check" button that opens a Winsock connection to see whether a server is online (default port 7171)
- English and Swedish UI, loaded from text files in `Languages/`
- minimize to the system tray

## Where it came from

I got into programming through Tibia. Around 5th grade, before I had ever used Linux, I got my own OpenTibia server running on that XP machine. I compiled it from C++ source so I could add in-game account creation: you could log in with any credentials, you spawned on a single tile in the void, you typed the account name and password you wanted into the chat, and then you logged in again with your new account. That was my first time building a C/C++ codebase, and I remember fighting the linker for a long time.

An IP changer was the other half of that setup. Players needed one to reach a server like mine, so I took one and made it my own. My server was called "New Age Network" (I didn't know what "New Age" meant yet), which is why this became the "New Age Ip Changer". The name later turned into "New Generation", shortened to NGenia, and my brother and I ran game servers together under that name (OpenTibia, a WoW private server, and eventually Minecraft) until 2013.

## Who wrote what

As far as I remember, this started from someone else's IP changer, and I made significant improvements on top of it. The project file still says `VersionCompanyName="Alex"`, a property that normally carries over from whoever created the project, so the original author was probably someone called Alex. I haven't been able to identify them or find the original online (2026 searches of GitHub and the Planet Source Code archive for its most distinctive lines turned up nothing). If this is yours, Alex, please get in touch.

Only one version of each file survives, so the split can't be shown commit by commit. From the file dates and the code:

- **Most likely mine (January 2007):** the language system (`mdlLanguage.bas`, `frmLanguage.frm`, `Languages/*.txt`, all saved within 15 minutes on 2007-01-27) and the Swedish translation; the About dialog; support for the then-new Tibia 7.91 client.
- **Most likely from the base I started with:** the core idea and structure. That covers the memory patching per client version, the loop that patches every open Tibia client, the server list, and the online check. I extended and reworked parts of it, but I can't say exactly which lines.
- **From neither of us:** `mdlTray.bas`. A comment in it calls it "the only part of the program that I haven't done myself"; the tray constants were copied from elsewhere.

The commits are authored by me because this is my copy of the code, not because I wrote every line in them.

## How it survived

The XP machine was never backed up, and the OpenTibia server itself is gone. This folder survived because I copied my old `PROG` folder from machine to machine. It ended up in Dropbox, went into my git-annex archive in 2017, and I found it again in 2026 while looking through old backups.

The commit dates are the file modification times preserved in that archive. The first commit contains only the two files that still carry 2006 dates; the 2006 versions of the other files were overwritten by my edits in January 2007. This README is the only file added afterwards.

## Notes

- `frmMain.log` is a VB6 load log (the Winsock control wasn't registered on some machine). `iplist.txt` is saved program state.
- Files are kept byte-for-byte as they were: CRLF line endings, and Windows-1252 encoding for `Languages/swedish.txt`.
- Building it needs Visual Basic 6 and `MSWINSCK.OCX`. The memory addresses only match those specific 2005–2007 Tibia clients, so it has no practical use today. It's here for the history.

Tibia is a trademark of CipSoft GmbH. This project has no affiliation with CipSoft.

---
*README written in 2026 with AI help, from the archived files and my own memories.*
