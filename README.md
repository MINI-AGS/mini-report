# Report Generator for the MINI test project

## How to build the project

The build of the project is done using the `pyinstaller` package. Note that before building the project, you need to have the virtual environment activated and the required packages installed.

````bash

The build is done using the following command depending on the operating system:

1. For Windows

```bash
pyinstaller src/main.py --onefile --name mini_report `--add-data="src/firebase-admin-sdk.json;." ` --hidden-import=openpyxl
````

2. For Linux

```bash
pyinstaller src/main.py --onefile --name mini_report --add-data=src/firebase-admin-sdk.json:.
```

## How to contribute

1. Clone the repository
2. Create a new branch
3. Make your changes
4. Push your changes to the repository to the branch you created
5. Create a pull request to the main branch
6. Contact Yeyo for the review of the pull request
7. Once the pull request is approved, merge the pull request to the main branch

## Installing a new package

1. Run the following command in the terminal

```bash
pip install <package-name>
```

2. Once the package is installed, run the following command in the terminal

```bash
pip freeze > requirements.txt
```

## How to run the code

1. Clone the Repository

2. Open the terminal and create a virtual environment

```bash
python3 -m venv venv
```

3. Activate the virtual environment

- For Windows

```bash
venv\Scripts\activate
```

- For MacOS/Linux

```bash
source venv/bin/activate
```

4. Run the following command in the terminal to install the required packages

```bash
pip install -r requirements.txt
```

5. Please contact Yeyo for the .env file with the required credentials

6. Once you have the .env file, run the following command in the terminal

```bash
python main.py
```

7. The code will run and generate the report
