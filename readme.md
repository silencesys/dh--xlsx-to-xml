# XLSX to XML

This little tool converts two column XLSX Excel files to XML.

<br>

## How to use

1. Install this package with NPM
```bash
npm install -g @silencesys/xlsx-to-xml
```
2. Create config file that will be used for the transformation
3. Run the tool
```bash
xlsx-to-xml --input your_input_file.xlsx --output your_output_file.xml --config your_config.json
```

<br>

### Available options
| Option | Description |
| ------ | ----------- |
| --input, -i | The input file to be converted to xlsx |
| --output, -o | The output file to be created |
| --config, -c | The config file that will be used |
| --dirty | Set this flag if you want to see XHTML tags added by XLSX transformer |
| --ignore-halves | Set this flag if you want to ignore rows that has one column empty |
| --help, -h | Show help |

<br>

## Config file
The config file is a standard JSON file with following structure:
```json
{
  "parentTagName": "dictionary",
  "rowTagName": "entry",
  "language": ["eng", "cze"],
  "stripTags": [
    "<span style=\"font-size:12pt;\">"
  ],
  "replaceTags": [
    {
      "from": "<span style=\"font-size:9pt;\">",
      "to": "<note>"
    },
  ],
  "divideBy": [
    [": "]
  ]
}
```
As you can see there are several options that can be used. **There is no default config file as each use case is expected to be unique.** You can use aforementioned snippet as your default config file.

| Option | Required | Description |
| ------ | :--: | ----------- |
| `parentTagName` | Yes | Name of the parent tag. |
| `rowTagName` | Yes | Name of the row tag. |
| `language` | Optional | List of languages that should be used. |
| `stripTags` | Optional | List of tags that should be stripped from the text. The list should contain only opening tags with all their attributes that should be removed. |
| `replaceTags` | Optional | List of tags that should be replaced. These tags should always be defined as a JSON object containing keys `from` and `to`. Only opening tags but with all attributes should be defined there. |
| `divideBy` | Optional | List of strings that should be used to divide the text. _You might want to include spaces following after these characters as the division method is quite dumb_.  |

<br>

### How to write a config file
I'm sure you ask how am I supposed to know which XHTML tags will be used in my XLSX document after conversion. Well, for that purpose there is a very simple way. Just run following command:
```bash
xlsx-to-xml --input your_input_file.xlsx --dirty
```
This will output a _dirty_ file containing XML with all XHTML tags. You can then decide which tags should be stripped or replaced and which should be kept. Just keep in mind that tags `<dirty-list>` and `<dirty-row>` are added by this tool and will be replaced by `parentTagName` and `rowTagName` respectively.

<br>

## Contributing
All the code is open source and you can contribute to the project by creating pull requests.

<br>

## License
This project is licensed under the [MIT license](LICENSE.md).
