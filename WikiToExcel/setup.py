'''
Created on Jul 28, 2015

@author: venkman69@yahoo.com
'''
from setuptools import setup
import subprocess
import os
# Fetch version from git tags, and write to version.py.
# Also, when git is not available (PyPi package), use stored version.py.
version_py = os.path.join(os.path.dirname(__file__), 'version.py')
try:

    version_git = subprocess.check_output(["git", "describe"]).rstrip()
except:
    with open(version_py, 'r') as fh:
        version_git = open(version_py).read().strip().split('=')[-1].replace('"','')

version_msg = "# Do not edit this file, pipeline versioning is governed by git tags"

with open(version_py, 'w') as fh:
    fh.write(version_msg + os.linesep + "__version__=" + version_git)

setup(name='wikitoexcel',
      version="{ver}".format(ver=version_git),
      description='Convert Wiki to Excel while maintaining formatting',
      url='http://github.com/venkman69/WikiToExcel',
      author='Narayan Natarajan',
      author_email='venkman69@yahoo.com',
      license='MIT',
      packages=['wikitoexcel'],
	  py_modules=['version'],
      install_requires=[
          'openpyxl',
          'beautifulsoup4'
      ],
      zip_safe=False)
