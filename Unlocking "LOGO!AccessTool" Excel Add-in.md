# How to unlock "LOGO!AccessTool" Excel Add-in
The Add-in comes in .xlam format, XLAM is an Macro-Enabled Add-In file that is used to add functions to spreadsheets. After installing the Add-in and opened the VBA Editor in Excel (alt-f11), we still wion't be able to access the source code, beacuse of a password protection that we don't know.\
![screenshot1](https://user-images.githubusercontent.com/43523843/160242869-e3a519f1-e4f2-4960-b0c1-a07fa91e4a0d.png)
Fortunately the Add-in file is just a compressed archive, so we can extract (unzip) all the underlying files, having done that we will see a structure as follows.

```bash
.
├── [Content_Types].xml
├── _rels
├── docProps
│   ├── app.xml
│   ├── core.xml
│   └── custom.xml
├── tree-md
└── xl
    ├── _rels
    │   └── workbook.xml.rels
    ├── printerSettings
    │   ├── printerSettings1.bin
    │   └── printerSettings2.bin
    ├── sharedStrings.xml
    ├── styles.xml
    ├── theme
    │   └── theme1.xml
    ├── vbaProject.bin
    ├── workbook.xml
    └── worksheets
        ├── _rels
        │   ├── sheet1.xml.rels
        │   └── sheet2.xml.rels
        ├── sheet1.xml
        └── sheet2.xml

8 directories, 17 files
```


