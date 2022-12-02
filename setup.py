import setuptools

with open("README.md", "r") as fh:
    long_description = fh.read()

setuptools.setup(
    name="ThinkcellBuilder",
    version="0.1.0",
    author="Philistino",
    author_email="philistino@protonmail.com",
    description="Simple unofficial Python library for generating PowerPoint presentations using Think-cell",
    long_description=long_description,
    long_description_content_type="text/markdown",
    keywords="automation powerpoint thinkcell ppttc business consulting",
    licence="MIT",
    url="https://github.com/Philistino/ThinkcellBuilder",
    packages=setuptools.find_packages(),
    classifiers=[
        "Programming Language :: Python :: 3.7",
        "Programming Language :: Python :: 3.8",
        "Programming Language :: Python :: 3.9",
        "Programming Language :: Python :: 3.10",
        "Programming Language :: Python :: 3.11",
        "License :: OSI Approved :: MIT License",
        "Operating System :: OS Independent",
    ],
    python_requires=">=3.7",
    project_urls={
        "Bug Reports": "https://github.com/Philistino/ThinkcellBuilder/issues",
        "Source": "https://github.com/Philistino/ThinkcellBuilder",
    },
)
