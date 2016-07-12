#!/usr/bin/env python3

from setuptools import setup

setup(name='taco.py',
      version='0.1a0',
      description='Simple converter of table-like data in YAML format to some common data types'
                  'like pdf or docx.',
      url='http://github.com/antonlukyanov/taco',
      author='Anton Lukyanov',
      author_email='anton.lukyanov@gmail.com',
      license='MIT',
      packages=['tableconv'],
      scripts=['taco.py'],
      zip_safe=False,
      include_package_data=True,
      install_requires=['python-docx>=0.8.6', 'pyyaml>=3.11', 'jinja2>=2.8'])
