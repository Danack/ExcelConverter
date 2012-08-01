How to build
============

Run 'ant' in the root directory of the project. 

One-JAR packages the application together with the dependency Jars into a single executable Jar file.


TODO list
=========

* Allow formatting to be preserved. Currently all formating is lost when the file is read e.g. a currency formatted cell 
that reads "$2.45" in Excel will be converted to "2.45".

* Implement float and double parsing of numbers.

* Check that all document formats work correctly - some may require libraries not present.