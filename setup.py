import pathlib
from setuptools import setup, find_packages

# The directory containing this file
HERE = pathlib.Path(__file__).parent

# The text of the README file
README = (HERE / "README.md").read_text()

print("Packages :",find_packages(exclude=("tests",)))
# This call to setup() does all the work
setup(
    name="AttachmentsExtractor",
    version="1.0.1",
    description="Extracts Embeded documents From Open XML documents. Such as excels, documents, presentations etc. .",
    long_description=README,
    long_description_content_type="text/markdown",
    url="https://github.com/surajForDD/AttachmentsExtractor",
    author="Suraj Kumar",
    author_email="surajkm09@gmail.com",
    license="MIT",
    classifiers=[
        "License :: OSI Approved :: MIT License",
        "Programming Language :: Python :: 3",
        "Programming Language :: Python :: 3.7",
    ],
    packages=find_packages(exclude=("tests",)),
    include_package_data=True,
    install_requires=["olefile"],
    entry_points={
        "console_scripts": [
            "AttachmentsExtractor=AttachmentsExtractor.__main__:main",
        ]
    },
)