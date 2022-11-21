================
google_dev_utils
================

A collection of classes to assist with accessing and manipulating Google services:

- Calendar
- Drive
- Sheets

NOTE: If you are running a version of Python < 3.7, a dependency deep in google's stack requires protobuf, which the current version of requires >3.7 python.

Instead, manually install the 3.19 version branch of protobuf before attempting to install this package:
```pip3 install protobuf==3.19.6```

Then this package will install cleanly.
