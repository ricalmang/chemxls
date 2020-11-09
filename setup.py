import setuptools
with open("chemxls\\supl\\README.txt", "r") as fh:
    long_description = fh.read()
setuptools.setup(
    name="chemxls",
    version="v0.0.1.post5",
    author="Ricardo Almir Angnes",
    author_email="ricardo_almir@hotmail.com",
    description="chemxls is a python & tk application for creating xls files (spreadsheets) bases on computational chemistry inputs and outputs.",
    long_description=long_description,
	license="MIT",
    url="https://github.com/ricalmang/chemxls",
	keywords = ['chemistry'],
	install_requires = ["xlwt","numpy"],
    packages=setuptools.find_packages(),
    classifiers=[
        "Programming Language :: Python :: 3",
        "License :: OSI Approved :: MIT License",
        "Operating System :: OS Independent",
    ],
    python_requires='>=3.8',
)