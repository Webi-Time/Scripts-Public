# PowerShell

## Collection de scripts PowerShell 
-------------------------------------

**This repository provides more than 500 useful and cross-platform PowerShell scripts in the [📂Scripts](Scripts/) subfolder - to be used by command-line interface (CLI), for remote control (via SSH), by context menu, by voice control (see repo [talk2windows](https://github.com/fleschutz/talk2windows)), by automation software (e.g. [Jenkins](https://www.jenkins.io/)), automatically on login/logoff/daily/etc., or simply to learn PowerShell.**

**[Download](https://github.com/fleschutz/PowerShell/releases) | [FAQ](Docs/FAQ.md)** | **Note:** the scripts support Unicode - a modern console is recommended (e.g. *Windows Terminal*)

### 🔊 Scripts d'audit
-----------------------------


####  🔊 Audit Active Diectory
| Script                                               | Description                                                                                     |
| ---------------------------------------------------- | ----------------------------------------------------------------------------------------------- |
| [audit-dnslog.ps1](Scripts/list-voices.ps1)          | Lists the installed text-to-speech voices. [Read more...](Docs/list-voices.md)                  |
| [list-voices.ps1](Scripts/list-voices.ps1)           | Lists the installed text-to-speech voices. [Read more...](Docs/list-voices.md)                  |
| [play-beep-sound.ps1](Scripts/play-beep-sound.ps1)   | Plays a short beep sound. [Read more...](Docs/play-beep-sound.md)                               |
| [play-files.ps1](Scripts/play-files.ps1)             | Plays the given audio files. [Read more...](Docs/play-files.md)                                 |

####  🔊 Audit Windows Server

####  🔊 Autres Audit
   
#### ⚙️ Scripts de configuration
| Script                                               | Description                                                                                     |
| ---------------------------------------------------- | ----------------------------------------------------------------------------------------------- |
| [list-voices.ps1](Scripts/list-voices.ps1)           | Lists the installed text-to-speech voices. [Read more...](Docs/list-voices.md)                  |
| [play-beep-sound.ps1](Scripts/play-beep-sound.ps1)   | Plays a short beep sound. [Read more...](Docs/play-beep-sound.md)                               |
| [play-files.ps1](Scripts/play-files.ps1)             | Plays the given audio files. [Read more...](Docs/play-files.md)                                 |



⚙️ Scripts de configuration
-------------------------------

| Script                                               | Description                                                                                       |
| ---------------------------------------------------- | ------------------------------------------------------------------------------------------------- |
| [add-firewall-rules.ps1](Scripts/add-firewall-rules.ps1) | Adds firewall rules for the given executables (needs admin rights). [Read more...](Docs/add-firewall-rules.md) |
| [check-cpu.ps1](Scripts/check-cpu.ps1)| Checks the CPU temperature. [Read more...](Docs/check-cpu.md)                                                    |
| [check-dns.ps1](Scripts/check-dns.ps1) | Checks the DNS resolution. [Read more...](Docs/check-dns.md)                                                    |
| [check-drive-space.ps1](Scripts/check-drive-space.ps1) | Checks a drive for free space left. [Read more...](Docs/check-drive-space.md)                   |
| [check-file-system.ps1](Scripts/check-file-system.ps1) | Checks the file system of a drive (needs admin rights). [Read more...](Docs/check-file-system.md)|
| [check-health.ps1](Scripts/check-health.ps1)         | Checks the system health. [Read more...](Docs/check-health.md)                                    |
| [check-ping.ps1](Scripts/check-ping.ps1)             | Checks the ping latency to the internet. [Read more...](Docs/check-ping.md)                       |
| [check-swap-space.ps1](Scripts/check-swap-space.ps1) | Checks the swap space for free space left. [Read more...](Docs/check-swap-space.md)               |


💻 Scripts for the Desktop
---------------------------

| Script                                               | Description                                                        | Help                                    |
| ---------------------------------------------------- | ------------------------------------------------------------------ | --------------------------------------- |
| [close-calculator.ps1](Scripts/close-calculator.ps1) | Closes the calculator application                                  | [Help](Docs/close-calculator.md)        |
| [close-cortana.ps1](Scripts/close-cortana.ps1)       | Closes Cortana                                                     | [Help](Docs/close-cortana.md)           |
| [close-chrome.ps1](Scripts/close-chrome.ps1)         | Closes the Chrome browser                                          | [Help](Docs/close-chrome.md)            |
| [close-program.ps1](Scripts/close-program.ps1)       | Closes the given program gracefully                                | [Help](Docs/close-program.md)           |


📁 Scripts for Files & Folders 
-------------------------------

| Script                                               | Description                                                        | Help                                |
| ---------------------------------------------------- | ------------------------------------------------------------------ | ----------------------------------- |
| [cd-autostart.ps1](Scripts/cd-autostart.ps1)         | Set the working directory to the user's autostart folder           | [Help](Docs/cd-autostart.md)        |
| [cd-desktop.ps1](Scripts/cd-desktop.ps1)             | Set the working directory to the user's desktop folder             | [Help](Docs/cd-desktop.md)          |
| [cd-docs.ps1](Scripts/cd-docs.ps1)                   | Set the working directory to the user's documents folder           | [Help](Docs/cd-docs.md)             |
| [cd-downloads.ps1](Scripts/cd-downloads.ps1)         | Set the working directory to the user's downloads folder           | [Help](Docs/cd-downloads.md)        |
| [cd-dropbox.ps1](Scripts/cd-dropbox.ps1)             | Set the working directory to the user's Dropbox folder             | [Help](Docs/cd-dropbox.md)          |
| [cd-home.ps1](Scripts/cd-home.ps1)                   | Set the working directory to the user's home folder                | [Help](Docs/cd-home.md)             |


♻️ Scripts to Convert Files
---------------------------

| Script                                                 | Description                                                                                         |
| ------------------------------------------------------ | --------------------------------------------------------------------------------------------------- |
| [convert-csv2txt.ps1](Scripts/convert-csv2txt.ps1)     | Converts a .CSV file to a text file. [Read more...](Docs/convert-csv2txt.md)                              |
| [convert-mysql2csv.ps1](Scripts/convert-mysql2csv.ps1) | Converts a MySQL database table to a .CSV file. [Read more...](Docs/convert-mysql2csv.md)                 |
| [convert-ps2bat.ps1](Scripts/convert-ps2bat.ps1)       | Converts a PowerShell script to a Batch script. [Read more...](Docs/convert-ps2bat.md)                    |
| [convert-ps2md.ps1](Scripts/convert-ps2md.ps1)         | Converts the comment-based help of a PowerShell script to Markdown. [Read more...](Docs/convert-ps2md.md) |
| [convert-sql2csv.ps1](Scripts/convert-sql2csv.ps1)     | Converts a SQL database table to a .CSV file. [Read more...](Docs/convert-sql2csv.md)                     |
| [convert-txt2wav.ps1](Scripts/convert-txt2wav.ps1)     | Converts text to a .WAV audio file. [Read more...](Docs/convert-txt2wav.md)                               |
| [export-to-manuals.ps1](Scripts/export-to-manuals.ps1) | Exports all scripts as manuals. [Read more...](Docs/export-to-manuals.md)                                 |


📝 Scripts for Git
-------------------

| Script                                               | Description                                                                                       |
| ---------------------------------------------------- | ------------------------------------------------------------------------------------------------- |
| [build-repo.ps1](Scripts/build-repo.ps1)             | Builds a Git repository. [Read more...](Docs/build-repo.md)                                       |
| [build-repos.ps1](Scripts/build-repos.ps1)           | Builds all Git repositories in a folder. [Read more...](Docs/build-repos.md)                      |
| [check-repo.ps1](Scripts/check-repo.ps1)             | Checks a Git repository. [Read more...](Docs/check-repo.md)                                       |
| [clean-repo.ps1](Scripts/clean-repo.ps1)             | Cleans a Git repository from untracked files. [Read more...](Docs/clean-repo.md)                  |


🔎 Scripts for PowerShell 
------------------------

| Script                                                 | Description                                                                                            |
| ------------------------------------------------------ | ------------------------------------------------------------------------------------------------------ |
| [daily-tasks.sh](Scripts/daily-tasks.sh)               | Execute PowerShell scripts automatically as daily tasks (Linux only). [Read more...](Docs/daily-tasks.sh.md) |
| [introduce-powershell.ps1](Scripts/introduce-powershell.ps1) | Introduces PowerShell to new users. [Read more...](Docs/introduce-powershell.md)                 |
| [list-aliases.ps1](Scripts/list-aliases.ps1)           | Lists all PowerShell aliases. [Read more...](Docs/list-aliases.md)                                     |
| [list-automatic-variables.ps1](Scripts/list-automatic-variables.ps1) | Lists the automatic variables of PowerShell. [Read more...](Docs/list-automatic-variables.md)|
| [list-cheat-sheet.ps1](Scripts/list-cheat-sheet.ps1)   | Lists the PowerShell cheat sheet. [Read more...](Docs/list-cheat-sheet.md)                             |


🛒 Various PowerShell Scripts 
------------------------------

| Script                                               | Description                                                        | Help                                    |
| ---------------------------------------------------- | ------------------------------------------------------------------ | --------------------------------------- |
| [add-memo.ps1](Scripts/add-memo.ps1)                 | Adds the given memo text to $HOME/Memos.csv                        | [Help](Docs/add-memo.md)                |
| [check-ipv4-address.ps1](Scripts/check-ipv4-address.ps1)| Checks the given IPv4 address for validity                      | [Help](Docs/check-ipv4-address.md)      |
| [check-ipv6-address.ps1](Scripts/check-ipv6-address.ps1)| Checks the given IPv6 address for validity                      | [Help](Docs/check-ipv6-address.md)      |
| [check-mac-address.ps1](Scripts/check-mac-address.ps1)| Checks the given MAC address for validity                         | [Help](Docs/check-mac-address.md)       |
| [check-subnet-mask.ps1](Scripts/check-subnet-mask.ps1)| Checks the given subnet mask for validity                         | [Help](Docs/check-subnet-mask.md)       |
| [check-weather.ps1](Scripts/check-weather.ps1)       | Checks the current weather for critical values                     | [Help](Docs/check-weather.md)           |
| [display-time.ps1](Scripts/display-time.ps1)         | Displays the current time for 10 seconds by default                | [Help](Docs/display-time.md)            |
| [list-anagrams.ps1](Scripts/list-anagrams.ps1)       | Lists all anagrams of the given word                               | [Help](Docs/list-anagrams.md)           |


Feedback
--------
Send your email feedback to: markus [at] fleschutz [dot] de

License & Copyright
-------------------
This open source project is licensed under the CC0 license. All trademarks are the property of their respective owners.


```
{
  "firstName": "John",
  "lastName": "Smith",
  "age": 25
}
```

Here's a sentence with a footnote. [^1]
foijflijfljl
ffdkl
----
dfljkfdlkjlkf
dfkjfdlkj

[^1]: This is the footnote.
- [x] Write the press release
- [ ] Update the website
- [ ] Contact the media
