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

* Force all strings to be converted to UTF-8 or strip bad characters with the equivalent of the PHP code:

function test( $text) {
    $regex = '/( [\x00-\x7F] | [\xC0-\xDF][\x80-\xBF] | [\xE0-\xEF][\x80-\xBF]{2} | [\xF0-\xF7][\x80-\xBF]{3} ) | ./x';
    return preg_replace($regex, '$1', $text);
}