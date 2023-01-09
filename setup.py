from setuptools import setup, find_packages

with open("requirements.txt") as f:
	install_requires = f.read().strip().split("\n")

# get version from __version__ variable in daily_report/__init__.py
from daily_report import __version__ as version

setup(
	name="daily_report",
	version=version,
	description="Daily Reports",
	author="abayomi.awosusi@sgatechsolutions.com",
	author_email="abayomi.awosusi@sgatechsolutions.com",
	packages=find_packages(),
	zip_safe=False,
	include_package_data=True,
	install_requires=install_requires
)
