Set-Content -Path "setup.py" -Value @"
from setuptools import setup, find_packages

setup(
    name="yosan",
    version="1.0.0",
    py_modules=["yosan", "auth", "server"],
    packages=find_packages(),
    install_requires=[
        "requests",
        "dnspython",
        "openpyxl",
        "reportlab",
        "libsql-client"
    ],
    entry_points={
        "console_scripts": [
            "yosan=yosan:main",
        ],
    },
)
"@
