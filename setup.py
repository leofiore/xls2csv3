from setuptools import setup, find_packages
from codecs import open
from os import path
import sys

here = path.abspath(path.dirname(__file__))

with open(path.join(here, 'README.rst'), encoding='utf-8') as f:
    long_description = f.read()

setup(
    name='xls2csv3',
    version='0.5.0',
    description='convert from Excel \'97 files into plain csv',
    long_description=long_description,
    url='',
    author='Leonardo',
    classifiers=[
        'Development Status :: 5 - Production/Stable',
        'Intended Audience :: Developers',
        'Programming Language :: Python :: 3',
        'Programming Language :: Python :: 3.3',
        'Programming Language :: Python :: 3.4',
        'Programming Language :: Python :: 3.5',
    ],

    py_modules=["xls2csv"],

    install_requires=[
        "xlrd==1.1.0",
    ],

    entry_points={
        'console_scripts': [
            'xls2csv=xls2csv:main',
        ],
    },
)

