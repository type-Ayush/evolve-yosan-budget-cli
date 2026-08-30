from setuptools import setup, find_packages

setup(
    name="yosan",
    version="1.0.0",
    description="Yosan Cloud Budget Management CLI",
    py_modules=["yosan", "auth", "server"],
    packages=find_packages(),
    install_requires=[
        "requests",
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
