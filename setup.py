from setuptools import setup, find_packages

setup(
    name="yosan",
    version="1.0.0",
    py_modules=["yosan", "auth"],
    install_requires=[
        "openpyxl>=3.1.0",
        "reportlab>=4.0.0",
        "requests>=2.28.0",
    ],
    entry_points={
        "console_scripts": [
            "yosan=yosan:main",
        ],
    },
)
