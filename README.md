# Text Converter

[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)
[![stability-stable](https://img.shields.io/badge/stability-stable-green.svg)](https://github.com/emersion/stability-badges#stable)

![preview](https://raw.githubusercontent.com/ManuelGil/text-converter/main/docs/images/preview.png)

This tool is a utility program that allows you to easily convert text.

## Features

- Convert Hexadecimal to Decimal.
- Convert Decimal to Hexadecimal.
- Convert ASCII/ANSI to Hexadecimal.
- Convert Hexadecimal to ASCII/ANSI.
- Convert ASCII/ANSI to Decimal.
- Convert Decimal to ASCII/ANSI.
- Convert ASCII/ANSI to Multy Byte UTF-8.
- Convert Multy Byte UTF-8 to ASCII/ANSI.
- Convert ASCII/ANSI to Java.
- Convert ASCII/ANSI to Visual Basic.

### Examples

- `Hexa2Dec`: FF FF FF => 255 255 255
- `Dec2Hexa`: 255 255 255 => FF FF FF
- `ANSI2Hexa`: foo => 66 6F 6F
- `Hexa2ANSI`: 66 6F 6F => foo
- `ANSI2Dec`: foo => 102 111 111
- `Dec2ANSI`: 102 111 111 => foo
- `ANSI2Java`: foo => \u0066 + \u006F + \u006F
- `ANSI2VB`: foo => Chr(&H66) + Chr(&H6F) + Chr(&H6F)

## Getting Started

This page will help you get started with Text Converter.

### Requirements

- Windows version compatible with Visual Basic 6.0
- Microsoft Visual Basic 6.0

### Installation

1. Clone or Download this repository
2. Unzip the archive if needed
3. Start Microsoft Visual Basic 6.0
4. Goto "File" > "Open project" > Search "TextConverter.vbp" file

## Built With

- Microsoft Visual Studio 6.0

## Authors

- **Manuel Gil** - _Owner_ - [ManuelGil](https://github.com/ManuelGil)

See also the list of [contributors](https://github.com/ManuelGil/text-converter/contributors)
who participated in this project.

## License

Text Converter is licensed under the MIT License - see the
[MIT License](https://opensource.org/licenses/MIT) for details.
