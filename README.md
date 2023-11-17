Please visit the following blog posts for knowing more about this project:

1) Short User Guide to creating Blogger Blogbooks from Backup/Export File using BlogBackupToBook VB.Net project and ExportFileFilterAndGenBook and another VBA projects' macros/code (free and open source), https://ravisiyermisc.blogspot.com/2023/09/short-user-guide-to-creating-blogger.html

2) Ported ExportFileFilterAndGenBook and ExportFileFilterByIndexList from VBA to BlogBackupToBook VB.Net single project of Visual Studio 2022 Community Edition, https://ravisiyermisc.blogspot.com/2023/09/ported-exportfilefilterandgenbook-from.html

Known issue in v4 release:
Cloning the current (which is v4 release) Github repo and then building the program works. But on running it to generate a blog book, an error message about BlogBook-Footer-Text.html not being found is shown. The blog book is still generated but without the BlogBook-Footer-Text.html content which is a minor issue.
A workaround to fix the problem is to copy BlogBook-Footer-Text.html from the main project source directory to the directory of the executable (bin\Debug\net6.0-windows or bin\Release\net6.0-windows on my setup based on whether I did a Debug build and/or a Release build). 
I do not want to try to fix this issue in the source code now. Don't know if I will fix it later on sometime. Of course, if somebody is interested to fix the issue and does it, and then shares the fix, it will be great!
