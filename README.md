# Reading UTF-8 files in Visual Basic 6
Reading UTF-8 files in Visual Basic 6


When reading text files encoded with UTF-8 with VB6 in a conventional way, the encoding is changed to ANSI, this causes several problems if the text has special characters and accents.

To solve this problem, we can perform the reading through the [**OpenTextFile**](https://docs.microsoft.com/pt-br/office/vba/language/reference/user-interface-help/opentextfile-method "OpenTextFile") method of the [**Scripting.FileSystemObjec**](https://docs.microsoft.com/pt-br/office/vba/language/reference/user-interface-help/filesystemobject-object "Scripting.FileSystemObjec") class, which allows the conversion to UTF-8, making the files with other encodings to be loaded correctly.
