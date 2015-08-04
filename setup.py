#! /usr/bin/env python
from distutils.core import setup
import sys

reload(sys).setdefaultencoding('Utf-8')

setup(
    name='xls_formatter',
    version='0.1',
    author='Piotrek Mońko',
    author_email='piotrek.monko@gmail.com',
    description='Utils for quick XLS responses',
    long_description=open('README.md').read(),
    url='https://github.com/piotrekmonko/xls_formatter',
    license='BSD License',
    platforms=['OS Independent'],
    packages=['xls_formatter'],
    include_package_data=True,
    classifiers=[
        'Development Status :: 3 - Alpha',
        'Environment :: Web Environment',
        'Framework :: Django',
        'Intended Audience :: Developers',
        'License :: OSI Approved :: BSD License',
        'Operating System :: OS Independent',
        'Programming Language :: Python',
        'Topic :: Software Development :: Libraries :: Python Modules',
    ],
)
