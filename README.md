# VSD

[![License][license-img]](LICENSE)
[![VSD18 Version][18ver]][18ver]
[![VSD19 Version][19ver]][19ver]
[![VSD20 Version][20ver]][20ver]

## About

The VSD is a series of spreadsheets created by L375 (Nanjou), which were used to track and analyze FRC match data at various events from 2018 through 2020.

The content in this repository is consireded archived and will not be updated except to fix critical issues.

All content in this repository is licensed under the terms of the [MIT license](LICENSE).

### VSD18 (POWER UP)

The VictoriaScoutingDatabase or VSD (now known as VSD18) is the first spreadsheet in the VSD series. It was created to analyze data from matches in the 2018 FRC season's game POWER UP. It is run solely using formulas, making it the least advanced of the VSDs. It is also quite laggy when not set to manually calculate, due to the overwhelming amount of array formulas in the `Data Processor` sheet.

### VSD19 (DESTINATION: DEEP SPACE)

VSD2019 (now known as VSD19) is the second spreadsheet in the VSD series, created to analyze data from matches in the 2019 FRC season's game DESTINATION: DEEP SPACE. VSD19 heavily improves upon its predecessor, introducing many new data tracking categories and making use of macros to eliminate the lag present in VSD18.

### VSD20 (INFINITE RECHARGE)

VSD20 is the third spreadsheet in the VSD series, created to analyze data from matches in the 2020/2021 FRC season's game INFINITE RECHARGE. VSD20 makes use of The Blue Alliance's API to pull team information and some match data from FIRST's database, while also streamlining the spreadsheet's configuration through the use of more macros and allowing for inputs in even more data tracking categories.

Dependencies (included in VSD20 download): [VBA-JSON](https://github.com/VBA-tools/VBA-JSON) and [VBA-Dictionary](https://github.com/VBA-tools/VBA-Dictionary)

## Setup

### VSD18

- Download the VSD18 spreadsheet
- \[add steps here]

### VSD19

- Download the VSD19 spreadsheet
- Enable macros when prompted
- Follow the steps under the `CONFIGURATION` section of the `README` sheet

### VSD20

- Download the VSD20 spreadsheet
- Enable macros when prompted
- Follow the steps under the `VSD Configuration` section of the `GUIDE` sheet

<!-- Labels -->
[license-img]: https://img.shields.io/github/license/Lowie375/VSD
[18ver]: https://img.shields.io/github/package-json/18ver/Lowie375/VSD?color=5B95F9&label=VSD18%20version
[19ver]: https://img.shields.io/github/package-json/19ver/Lowie375/VSD?color=7F4EC8&label=VSD19%20version
[20ver]: https://img.shields.io/github/package-json/20ver/Lowie375/VSD?color=E32D91&label=VSD20%20version
