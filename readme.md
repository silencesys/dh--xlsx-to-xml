# XLSX to XML

This little tool converts two column XLSX Excel files to XML.

## How to use

1. Install this package with NPM
```bash
npm install -g @silencesys/xlsx-to-xml
```
2. Create config file that should be used for the transformation
3. Run the tool
```bash
xlsx-to-xml --input your_input_file.xml --output your_output_file.xml --config your_config.json
```

### Available options
| Option | Description |
| ------ | ----------- |
| --input, -i | Set path to input file |
| --output, -o | Set output file path |
| --config, -c | Set path to config file |
| --help, -h | Show help |

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

## Contributing
All the code is open source and you can contribute to the project by creating pull requests.

## License
This project is licensed under the [MIT license](LICENSE.md).
