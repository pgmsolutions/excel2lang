# excel2lang

excel2lang is a small utility program which takes a xlsx file as an input and extracts the languages strings to a JSON format.
It has several cool features:
- Support comment lines
- Can create multiple files
- Use the first column as default if no value

## Installation

```
npm install -g excel2lang
```

## Usage

```
excel2lang input.xlsx
```

## xlsx format

Your excel must be like this:

| #   | Unique ID            | English | French       |
| --- | -------------------- | ------- | ------------ |
| C   | id                   | en      | fr           |
| F   | ./{lang}.json        |         |              |
| #   | Generic              |         |              |
|     | open                 | Open    | Ouvrir       |
|     | browse               | Browse  | Parcourir    |
|     | edit                 | Edit    | Modifier     |
|     | select               | Select  | Selectionner |
| F   | ./topbar/{lang}.json |         |              |
| #   | Topbar               |         |              |
|     | topbar/file          | File    | Fichier      |
|     | topbar/help          | Help    | Aide         |

The first column indicates what the line is for:

| Character | Description                                                                                                                                                         |
| --------- | ------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| _nothing_ | Regular line of strings.                                                                                                                                            |
| C         | Code line, indicating the language code for each column.                                                                                                            |
| F         | The JSON file to create. Will include all lines until the end of the file or another _F_ line. The `{lang}` substring will be replace by the current language code. |
| #         | Comment line.                                                                                                                                                       |

In order to work, a _C_ line __MUST__ be the first line of the excel sheet (beside comment line). Also, a _F_ line __MUST__ be present before the first string line.

## TODO
- More checks for errors