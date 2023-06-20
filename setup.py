from setuptools import setup
import codecs
import os

here = os.path.abspath(os.path.dirname(__file__))

with codecs.open(os.path.join(here, "README.md"), encoding="utf-8") as fh:
    LONG_DESCRIPTION = "\n" + fh.read()

DESCRIPTION = 'Control an app without disturbing your everyday life.'

# Setting up
setup(
	name="pyappauto",
	url="https://github.com/User00092/PyAppAuto",
	version='1.0.2',
	license='MIT',
	author="User0092",
	author_email="unknownuser0092@protonmail.com",
	description=DESCRIPTION,
	long_description_content_type="text/markdown",
    long_description=LONG_DESCRIPTION,
	keywords=['Windows App control', 'App controller', 'PyAppAuto'],
	python_requires='>=3.7.0',
	install_requires=['pywin32==306', 'opencv-python==4.7.0.72', 'pygetwindow==0.0.9', 'numpy==1.24.3'],
	classifiers=[
		'Development Status :: 5 - Production/Stable',
		'Topic :: Utilities',
		'Operating System :: Microsoft :: Windows',
		'Programming Language :: Python',
		'Intended Audience :: Developers'
	]
)