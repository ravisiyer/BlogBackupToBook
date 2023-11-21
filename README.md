BlogBackupToBook software creates HTML blogbook(s) from Blogger blog XML Backup/Export files. It also allows for filtering the output HTML blogbook(s) by content matching strings, by published and/or updated date range and by an index(es) list where the index(es) are positions of posts and pages entries in the Blogger XML Backup/Export file. 

As compared to release v3, the latest release v4 has undergone more tests but it has not been thoroughly tested. I do not have the time now to do that thorough testing and bug-fixing as I have got busy with other stuff. This is the release interested persons can try out and see if it meets their needs.

If I recall correctly, the setup file program in v3 release had some issues which were not easily getting resolved. The setup file part has been dropped in this release v4 as I don't have time for refreshing where I was on it, and then taking it forward. This current v4 release does not have setup file folder and project but earlier v3 release has it. No other changes have been made to v3 release in v4 release i.e. the main project - BlogBackupToBook - which is mapped to folder VB-ExportFileFilterAndGenBook (as that was the earlier project name) - is unchanged. No changes have been made to any of its files including source files.

If somebody wants to fix the setup file issues and then use it, he/she can use v3 release.

For more details, visit:

1) Short User Guide to creating Blogger Blogbooks from Backup/Export File using BlogBackupToBook VB.Net project and ExportFileFilterAndGenBook and another VBA projects' macros/code (free and open source), https://ravisiyermisc.blogspot.com/2023/09/short-user-guide-to-creating-blogger.html

2) Ported ExportFileFilterAndGenBook and ExportFileFilterByIndexList from VBA to BlogBackupToBook VB.Net single project of Visual Studio 2022 Community Edition, https://ravisiyermisc.blogspot.com/2023/09/ported-exportfilefilterandgenbook-from.html

Known issue in v4 release:
Cloning the current (which is v4 release) Github repo and then building the program works. But on running it to generate a blog book, an error message about BlogBook-Footer-Text.html not being found is shown. The blog book is still generated but without the BlogBook-Footer-Text.html content which is a minor issue.

A workaround to fix the problem is to copy BlogBook-Footer-Text.html from the main project source directory to the directory of the executable (bin\Debug\net6.0-windows or bin\Release\net6.0-windows on my setup based on whether I did a Debug build and/or a Release build). 

I do not want to try to fix this issue in the source code now. Don't know if I will fix it later on sometime. Of course, if somebody is interested to fix the issue and does it, and then shares the fix, it will be great!
