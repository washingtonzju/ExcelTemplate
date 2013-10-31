ExcelTemplate
=============

A Friendly API for Generating Excel Report by Java.

These APIs save time for data formats, and programmers could focus on the data generation.

Dependencies:

This libarary is based on Apache POI. To make the sample work, you need to add:

######poi-3.9-20121203.jar
######poi-ooxml-3.9-20121203.jar
######poi-ooxml-schemas-3.9-20121203.jar
######ooxml-lib\dom4j-1.6.1.jar
######ooxml-lib\stax-api-1.0.1.jar
######ooxml-lib\xmlbeans-2.3.0.jar

to your library.

Formats:

Data is orginized into a List < List < Object > > data set.

You could specify the background color, font color and margin for each sheet.

Titles attributes are seperated with datas'.

Extendtion is welcome, you could wrote subclass to display in zebora lines, or columns, etc.

It's also a good example that shows how to use POI

