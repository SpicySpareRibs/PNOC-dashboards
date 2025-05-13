# PNOC-dashboards
This repository contains the recompiler application developed in Python to be used on the datafiles provided by PNOC.

## Setup

This application is managed using Poetry. To setup the project and install the dependencies, run the following command after cloning the repository
```
poetry install
```

## Building the Application

The following command builds application.py into a single .exe file (the -F option) that does not run a console (the -w option) since the recompiler application runs fully through its GUI. The built .exe file can be shared and works as a standalone file.

```
poetry run pyinstaller -F -w application.py
```
