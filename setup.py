from setuptools import setup

setup(
    name="yosan",
    version="1.0.1",
    py_modules=["yosan", "auth", "server"],
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
