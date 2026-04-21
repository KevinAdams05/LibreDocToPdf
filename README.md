[Release Notes](ReleaseNotes.md)

This is a simple program to convert DOC and DOCX files to PDFs. It uses LibreOffice and requires LibreOffice to be installed.

The Windows version is written in C# and is a .NET8 WinForms program.

The Linux version is written in Vala and GTK.

**Usage**
* Select a folder path by either using the Browse button or dragging and dropping the folder into the path text box
* Press convert
* The app will then look for all DOC or DOCX files in that folder and convert them to PDF
* There are settings for: number or retries, output log to a file, and setting output folder to be something other than the source

**Windows version**

Tested on Windows 11, but should work on most versions of Windows that support LibreOffice.
<img width="519" height="382" alt="Screenshot" src="Screenshot - Windows.png" />


**Linux version**

Tested on Linux Mint 22.3 Cinnamon.

<img width="519" height="452" alt="vala screenshot" src="https://github.com/user-attachments/assets/582a5a32-0a12-4b3f-9055-a8020ce5bcb4" />


**Notes**

This is meant to be a quick and simple app, there may be bugs. Feel free to log an issue or submit a PR.

Some AI was used to help, especially with the Linux port.

**License**

This project is licensed under the [MIT License](LICENSE).
