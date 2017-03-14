from setuptools import setup

setup(
    name='xlwrap',
    version='0.1.2',
    description='A library to manipulate excel files of multiple types with an identical, simple API',
    url='https://github.com/erik-sn/xlwrap.git',
    author='Erik Niehaus',
    author_email='nieh.erik@gmail.com',
    license='MIT',
    classifiers=[
        'Development Status :: 3 - Alpha',
        'Intended Audience :: Developers',
        'License :: OSI Approved :: MIT License',
        'Programming Language :: Python :: 3',
    ],
    keywords='excel development',
    py_modules=["xlwrap"],
    install_requires=['xlrd', 'openpyxl'],
)
