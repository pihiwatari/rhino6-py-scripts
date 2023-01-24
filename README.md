# Rhino6 custom command scripts

A collection of custom rhino commands and functionalities used for tooling development, file management and geometry export.

#### List of commands

- ExportToCNC: Exports files in STEP format, using custom directives for filename and location.
- GetBoundingBoxDimensions: Uses BoundingBox command to return only the dimensions of the selected object(s) in 'x _ y _ z' format.

#### Auxiliary files

- csv_crud.py: Used in the ExportToCNC command. Creates new files, adds and updates data to csv files, this information is then uploaded to a master excel control file.

## Installation

Double click on the **_'ModuleName {GUID}.rhi'_** file you want to install and follow installation instructions. This would copy the content of this plugin in the corresponding folders automatically.
Alternatively you can download this repo, then go to your **'_%APPDATA%\McNeel\Rhinoceros\5.0\Plug-ins\PythonPlugIns_'** extract the files using the extract to **_'FolderName {GUID}\\'_** option using your prefered zip software.

**_Note:_** Rhino should be closed when “installing” new plugins, otherwise it will need to be restarted before it recognizes any new commands.

## Using this module

After installing just type in the command line the name of the command as it appears on the **_'CommandName_cmd.py_** file. If the command doesn't appear in the command list, then execute the **_'EditPythonScript'_** command to load the python environment, the commands should appear now as you type them.
