TFS Release Notes
=================

TFS Release Notes automatically generates release notes for a project hosted on Team Foundation Server using MS Word or MS Excel.
The release notes contain all PBIs in a given iteration plus any bugs in that iteration with specific tags. 
They also contain the test cases for those work items.

Prerequisites
------------------
* TFS Release Notes generates MS Word and MS Excel documents using their respective interops, so you
must have Word/Excel installed on the machine you will use to generate the release notes.
* The Windows user running ReleaseNotes.exe must have read access to the team project specified.

Usage
------------------
Run the executable (ReleaseNotes.exe) with exactly eight arguments, in the following order: 

* TFS Project Collection Path (eg https://mytfsserver.com/MyProjectCollection)
* TFS Project Name (eg MyProject)
* Iteration (The full name of the sprint/iteration for which to generate release notes)
* Release Notes Type (one of  `excel`, `word`, `html`)
* Database Server (arbitrary string - the name of the database server)
* Web Server (arbitrary string - the name of the web server)
* Database (arbitrary string - the name of the database)
* Web URL (arbitrary string - the location of the published application)