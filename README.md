Groovy UNO Extension
====================

Introduction
------------
This is an Apache Groovy language extension of the Java UNO API's. The artifact of this extension is a Java jar file 
that when used in a Groovy script or class adds convenience methods to the regular Java UNO API's.

The goal of the Groovy UNO Extension is to allow UNO programming that is less verbose than using the Java UNO API's alone.

These methods are implemented using Groovy Extension Modules. An extension module allows you to add new methods to 
existing classes, including classes which are precompiled, like classes from the JDK or in this case Java UNO classes. 
These new methods, unlike those defined through a metaclass or using a category, are available globally.

Aside from a few general methods, initial efforts have been on enhancing the spreadsheet API's and future work will be 
on enhancing the other applications. 

Background
------
This project's home was originally the Apache OpenOffice Subversion repository in the developer tools area and 
original documentation is on the OpenOffice wiki [Groovy UNO Extension] [1].

Current development work by the author and issue tracking are now here and documentation is on the [GUNO Documentation page][2].

Building
--------
This software uses the Gradle build system and the library jar can be built with:
> ./gradlew jar

The jar task will build the library and groovydoc jar files.

For a complete build sutable for a Maven repository with library jar, groovydoc jar, and source jar including checksums and signing.
> ./gradlew publish

To build and install into your local Maven cache.
> ./gradlew publishToMavenLocal

For the signing tasks you will need a gradle.properties file containing:
> signing.keyId=your key id 
> signing.password=your password
> signing.secretKeyRingFile=path to your secret keyring file

Remember to exclude this file from your repository.

License
-------
This software is licensed under the Apache License 2.0

See the LICENSE, and NOTICE files for more information.

[1]: https://wiki.openoffice.org/wiki/Groovy_UNO_Extension "Groovy UNO Extension"
[2]: https://cbmarcum.github.io/guno-extension/ "GUNO Documentation"
