# Unified schedule
This is a parser that unifies schedules for Software Engineering specialty and Faculty of Economics of NaUKMA

# Table of contents
1. [How to use](#how-to-use)
2. [Technologies used](#technologies-used)
3. [Output format](#output-format)
# How to use
Clone a repository to your local machine and switch to project directory:
```
git clone https://github.com/TeksPeks/unified_schedule.git
```
```
cd unified_schedule
```
Install dependencies:
```
pip install -r requirements.txt
```
Run the script:
```
python ./main.py
```
The output will be saved in `output.json` file.

# Technologies used
- Python
  - [`openpyxl`](https://pypi.org/project/openpyxl/) - for reading excel sheets
  - `re` - for working with strings
  - `json` - for saving output as json

# Output format
Output is saved as `json` in `output.json` file with structure similar to an example with little changes:
```
{
	"Факультет": {
		"Назва спецільності": {
			"Назва предмету": {
				"групи": {
					"час": "",
					"тижні": "номери тижнів, коли йде цей предмет",
					"аудиторія": "",
					"день тижня": ""
				}
			}
		}
	}
}

```
