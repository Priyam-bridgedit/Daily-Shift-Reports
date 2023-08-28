import sys
from cx_Freeze import setup, Executable

# Name of your main script
target_script = 'shirftReport.py'

# Build options
build_options = {
    'packages': ['os', 'tkinter', 'pandas', 'pyodbc', 'configparser'],
    'excludes': [],
    'include_files': ['config.ini'],  # Include the config.ini file
}

# Executable options
executables = [
    Executable(
        script=target_script,  # The name of your main script
        base=None,  # Use the default base (Windows) for the executable
        targetName='Shift Report.exe',  # The name of the executable
    )
]

# Create the setup
setup(
    name='Shift Report',  # Name of the application
    version='1.0',
    description='Shift Report',
    options={'build_exe': build_options},
    executables=executables,
)
