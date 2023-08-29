# Message-Template-Generation

## About

This testing helper framework is built to be as simple as possible to maintain:

- **Templates**: Are stored in the */templates* folder where tags are denoted by the following pattern: $[tag_name] and saved as a basic txt file.
- **Data**: Test case data is stored in the */data* folder as a simple Excel file. The file must have the same name as the template file name but as a .xlsx file. Each tag used in the template file should have a corresponding header value in the Excel file in row 1 with n number of message content saved in subsequent rows.
- **Output**: The output messages will be stored in the */output* folder. Each time the application is run it clears the output folder as to not contaminate the data that can be uploaded to IMS.

Once the application has generated an output, the user will be asked if they wish to send the messages directly to IMS for processing.

## Helpful resources

- [Python templating](https://wiki.python.org/moin/Templating)
- [pysftp](https://pysftp.readthedocs.io/en/release_0.2.9/)

## Python PIP list

Please find, below, the list of python packages that were used at the time of building this framework:

|Package            |Version |
|-------------------|--------|
|pysftp             |0.2.9   |

*It is strongly recommended to create a Python Virtual Environment and get the abovementoned packages installed on this environment.*

### Command to create a new Python virtual environment

```bash
python -m venv /path/to/new/virtual/environment
```

### Command to install the packages when using Umbrella

```bash
pip install --trusted-host pypi.org --trusted-host files.pythonhosted.org pip pysftp
```

## Python version

python --version:

**Python 3.11.4**

## Sample BAT script to automate running the tests from your desktop

```bash
@echo off
cd C:\PycharmProjects\VSCodeProjects\MessageTemplateGeneration\message-template-generation
:start
python main.py
goto start
```
