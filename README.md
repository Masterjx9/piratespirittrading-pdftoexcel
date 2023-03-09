# How to use pdf to excel for pirate spirit trading

## Prequists 
1. Python https://www.python.org/downloads/
2. (Optional but extremely helpful) Git https://git-scm.com/download
3. (Optional but required) VSCode https://code.visualstudio.com/Download

## Install setup
1. After installing python and git perform the following command: 
<br>

```git clone https://github.com/Masterjx9/piratespirittrading-pdftoexcel.git```

2. Then use `cd piratespirittrading-pdftoexcel` to change to the directory you just downloaded from github.

3. Then perform the following command to create a virtual environment (Allowing you to download python packages for JUST this project. This is useful so if you have other python projects, it would overwrite the packages you want for those projects) <br>
```python -m venv .venv``` <br>
or <br>
```python3 -m venv .venv``` <br>
or <br>
```py -m venv .venv``` <br>
NOTE: This depends on how your python is installed and what OS you are using. <br>
Then perform this command to activate your virtual environment: <br>
```.venv/scripts/activate``` <br>
or <br>
```source .venv/bin/activate``` <br>
NOTE: This depends on if you are on windows or linux/mac

4. Then perform the follow command to install the required packages: <br>
```pip install -r requirements.txt```

## How to use script/app
At the command line type in the following: <br>
```python pdftoexcel.py <ENTER YOUR PDF PATH HERE>``` <br>
<br>
Example: <br>

```python pdftoexcel.py "3-7-2023\line sheet - stock report-SW 3723.pdf"``` <br> or <br>
```python pdftoexcel.py "line sheet - stock report-SW 3723.pdf"```
