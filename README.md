# GREATEST Unit-Test Runner

## Introduction
This is a small windows application to call the `SUITE`'s available in the Test-Runner, created by the [GREATEST Unit-Test Framework](https://github.com/silentbicycle/greatest). This program expects a dropped *.exe file (or per first command line parameter) to work with.

## Getting Started
Just download the [executable](https://github.com/Taneeda/gutr/blob/master/greatestUnitTestRunner.exe) and start it. After start, the program await for a test runner to drop. The dropped file will be executed with the AutoIt [Run](https://www.autoitscript.com/autoit3/docs/functions/Run.htm)-Function and the `-h` parameter, to check whether the dropped file is an GREATEST Test-Runner.

![DragAndDrop](https://raw.githubusercontent.com/Taneeda/gutr/master/img/dragDrop.png)
![LoadedSUITEs](https://raw.githubusercontent.com/Taneeda/gutr/master/img/SUITEsLoaded.png)
![ReportExample](https://raw.githubusercontent.com/Taneeda/gutr/master/img/reportExample.png)

## Compile yourself
Simply download and install [AutoIt](https://www.autoitscript.com/site/) and compile the [script-file](https://github.com/Taneeda/gutr/blob/master/greatestUnitTestRunner.au3) with it.
