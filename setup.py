from setuptools import setup
import pkg_resources

packages = [
    'pygs'
]

install_requires = [
    'google-api-python-client>=1.6.2',
    'pandas>=0.20.1',
    'oauth2client'
]

long_desc = """This allows a user to send a dataframe to a Google Sheet"""

version = '0.9.6'

setup(
    name="pygs",
    version=version,
    description="Pandas DataFrames to Google Sheets",
    long_description=long_desc,
    author="JP Schultz",
    author_email="jp.schultz@gmail.com",
    install_requires=install_requires,
    packages=packages,
    package_data={}
)
