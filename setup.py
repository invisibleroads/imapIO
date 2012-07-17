'Setup script for imapIO'
import os

from setuptools import setup, find_packages


here = os.path.abspath(os.path.dirname(__file__))
README = open(os.path.join(here, 'README.rst')).read()
CHANGES = open(os.path.join(here, 'CHANGES.rst')).read()


setup(
    name='imapIO',
    version='0.9.5',
    description='Convenience classes and methods for processing IMAP mailboxes',
    long_description=README + '\n\n' +  CHANGES,
    license='MIT',
    classifiers=[
        'Intended Audience :: Developers',
        'Programming Language :: Python',
        'License :: OSI Approved :: MIT License',
    ],
    keywords='imap',
    author='Roy Hyunjin Han',
    author_email='service@invisibleroads.com',
    url='https://github.com/invisibleroads/imapIO',
    install_requires=['chardet'],
    packages=find_packages(),
    include_package_data=True,
    test_suite='imapIO.tests',
    tests_require=['nose'],
    zip_safe=True)
