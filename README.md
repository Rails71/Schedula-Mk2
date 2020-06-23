# Schedula Mk2

An excel-python combination to make appointing to officials easier.

For the advanced user there is a python API for interfacing with [schedula](https://schedula.sportstg.com/).

## Getting Started

These instructions will get a copy of the tool up and running.
There are two parts to the tool, a command-line utility and an excel spreadsheet.

The good news is if you are using Windows it may just work (depending on your security settings). On a mac you will need to install python (see below) and compile the command line tool.

### Windows

1. Open "Schedula Mk2.xlsm". Be sure to enable macros.
2. Navigate to the controll panel sheet.
3. (Optional) Enter a schedula username in the approopriate cell.
4. Press "Update Official List", enter password when prompted. Wait for the shell to close.
5. Enter the start day and the number of days of fixtures to get.
6. Press "Pull Some". Wait for the shell to close.
7. Press "Load to Sheet".
8. After entering appointments on the Schedula sheet, press "Push Appointments".


### Mac - comming soon (maybe)

## Development

### Prerequisites

The following software is required:

* Microsoft Excel
* Python 3.7+

Excel provides a neat GUI wrapped arround the command-line tool. The command line tool can be used without Excel (with difficulty).

### Installing Prerequisites

#### Excel

Installing Excel is left to the user

#### Python

Install python - available [here](https://www.python.org/downloads/)

The requests library is required. It can be installed with pip

```
pip install requests
```

## Built With

* [Python 3.x](https://www.python.org/) - Backend interface with [schedula](https://schedula.sportstg.com/)
* [Microsoft Excel](https://www.microsoft.com/en-au/microsoft-365/excel) - User Interface
* [Visual Basic for Applications](https://docs.microsoft.com/en-us/office/vba/api/overview/) - Excel Magic

## Contributing

Contributions welcome, just open an issue or a pull request. 

## Authors

See the list of [contributors](https://github.com/Rails71/schedulaMk2/graphs/contributors).

## License

This project is licensed under the [MIT License](https://github.com/Rails71/schedulaMk2/blob/master/LICENCE.txt).
