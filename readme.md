Synopsis

NpJson.cls - json parser for vb6/vba

Examples

Open file JSON.xlsm from directory .\src. You should have a dialog window. Enabled visual basic macros is it required.
Dialog windows contains follow elements:
  * File combobox - list of files *.js from .\scr\data
  * Refresh - refresh list of files in combobox
  * Extended syntax checkbox - some extended feaches enavled (object keys without quoted, more fluent syntax on numbers, ...)
  * Load as script check box - load javascript variables definitions to json (see example .\scr\data\script.js)
  * Load File button - load selected file with selected options.
  * JSON tree - view JSON when it successfully loaded; nodes of tree are contains JSON values and type in inside of # pair.
  * Save to Data\Tmp.js button - save file to file. May be used to check valid JSON saving
  * Close button - close form; press Yes on close query when you want to close Excel

Motivation

Parser was nessecery in one of project. Workable parser was not found, so, now it exists.

Installation

Put NpJson.cls to you project. Use JSON.xlsm as example to use.

Tests / Debug

If you want to test or debug parser, please put you JSON to file with extention .js to .\scr\data. Open JSON.xlsm and load you JSON.

Contributors

Sergey Mogilnikov, sergey.mogilnikov@kpbs.ru

License

The MIT License (MIT)
