# Accounts

## Requirements

- [Python]("https://www.python.org/") 3.9 or greater
- [Python Poetry]("https://python-poetry.org/docs/") 1.1 or greater
- Accounts Excel file with correct column mapping
- S-26 PDF form

## Installation

Steps:

- Clone the project
- Set up your environment (.env)
- Install the requirements with poetry (no root)
- Add your S-26 file to the sources folder

```powershell
git clone https://github.com/rr0001/accounts.git
cd accounts
cp .env.dist .env
poetry install --no-root
```

## Usage

Once you have set up the project and the environment, all you need to do is run the following command:

```powershell
py go.py
```

The script will use the PDF source file to create a PDF for each sheet in the source Excel file. The files will be named based on the sheet name. The files will be saved in the OUT folder.
