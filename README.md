html2oft
========

**KNOWN ISSUES**
* Doesn't work. See: http://blog.rdkl.us/follow-up-generate-oft-files-from-html-with-mono-net/

Mono.NET project using OpenMCDF to inject HTML into a template OFT file. Meant to not require Windows.

See my blog entry for proper credits but the big two are:

* http://stackoverflow.com/questions/7957827/is-there-a-difference-between-the-outlook-msg-and-oft-file-formats
* http://sourceforge.net/projects/openmcdf/

Blog entry: http://blog.rdkl.us/generate-oft-files-from-html-with-mono-net/

Usage: html2oft infile.html outfile.oft

Assumes a blank OFT exists in the current directory named Blank.oft.
