# TIA Openness Excel Importer

1. Start TIA Portal and open a project including a WinCC Unified device
2. Run the Excel Exporter. Excel files will be created, containing the screen information.
3. Adjust the Excel files
4. Run the Excel Importer. The screen items will be adjusted. If the screen is deleted before or does not exist, it will be created.

## Syntax
ExcelExporter.exe \<hmiruntimename> [-h] [-id=\<processid>] [-i="\<screennames>" | -e="\<screennames>"] [-da="\<definedattributes>"] [-p="\<projectpath>"] [-ui=yes|no] [-c=yes|no] [--l=Debug|Info|Warning|Error]

## Parameter

| short | long | details | description |
| ----- | ---- | ------- | ----------- |
|-h | --help | IsReqired: False | if this flag is set, help will be shown and application will be closed afterwards, ignoring  all other options |
|-id  |   --processid    |    (default: ) IsReqired: False      |       define a process id the tool connects to. If empty, the first TIA Portal process will be connected to|
|-i   |   --include      |     (default: ) IsReqired: False     |       add a list of screen names on which the tool will work on, split by semicolon (cannot be combined with --exclude), e.g. "Screen_1;My screen 2"|
|-e   |   --exclude      |      (default: ) IsReqired: False    |       add a list of screen names on which the tool will not work on, split by semicolon (cannot be combined with --include), e.g. "Screen_1;My screen 2"|
|-da|--definedattributes | (default: ) IsReqired: False| define a list of attributes/properties that will be exported only (but Name and Type is always included) to speed up exporting and focus only on the relevant properties, e.g. "Left;Top;Font.Size" |
|-p   |   --projectpath  |       (default: ) IsReqired: False   |       if you have no TIA Portal opened, the tool can open it for you and open the project from this path (ProcessId will be ignored, if this is set), e.g. D:\projects\Project1\Project1.ap18 |
|-ui  |   --showui       |    (default: yes) IsReqired: False   |          if you provided a ProjectPath via -p you may decide, if TIA Portal should be opened with GUI or without, e.g. "yes" or "no"|
|-c   |   --closeonexit  |       (default: no) IsReqired: False |         you may decide, if the TIA Portal should be saved and closed when this tool is finished, e.g. "yes" or "no"|
|-l   |   --loglevel     |     (default: Info) IsReqired: False |           define a log level: Debug,Info,Warning,Error|

## Examples
- Export 2 screens: ExcelExporter.exe HMI_RT_1 --include="Screen_3;Screen_4"
- In case of errors, additional logs can be produced: ExcelExporter.exe HMI_RT_1 --loglevel=Debug
- Export only 3 properties: ExcelExporter.exe HMI_RT_1 -da="Left;Top;Font.Size"
- Full silent - Start TIA Portal and project, export and close TIA Portal: ExcelExporter.exe HMI_RT_1 -p="D:\\MyProject\\MyProject.ap20" -c=yes
