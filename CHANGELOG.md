Yii 2 Spreadsheet Data Export extension Change Log
==================================================

1.0.3, January 24, 2019
-----------------------

- Bug #11: Fixed `Spreadsheet::send()` fails for 'Xlsx' writer (klimov-paul)


1.0.2, April 18, 2018
---------------------

- Bug #4: Fixed `Spreadsheet::save()` forces directory name to lowercase (klimov-paul)
- Bug #5: Fixed `Spreadsheet::createDataColumn()` sets column format to 'text' instead of 'raw' (klimov-paul)


1.0.1, April 9, 2018
--------------------

- Enh #2: `Spreadsheet::send()` now throws `\RuntimeException` in case temporary file can not be created (Eseperio, klimov-paul)


1.0.0, February 13, 2018
------------------------

- Initial release.
