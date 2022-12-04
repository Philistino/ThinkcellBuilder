# ThinkcellBuilder 📊
[![CI](https://github.com/philistino/thinkcellbuilder/actions/workflows/main.yml/badge.svg)](https://github.com/philistino/thinkcellbuilder/actions/workflows/main.yml) [![codecov](https://codecov.io/gh/philistino/thinkcellbuilder/branch/main/graph/badge.svg?token=F71I6S66YW)](https://codecov.io/gh/philistino/thinkcellbuilder) [![PyPI version shields.io](https://img.shields.io/pypi/v/thinkcell.svg)](https://pypi.python.org/pypi/thinkcellbuilder/) [![Supported Python versions](https://img.shields.io/pypi/pyversions/thinkcellbuilder.svg)](https://pypi.org/project/thinkcellbuilder/) [![Code style: black](https://img.shields.io/badge/code%20style-black-000000.svg)](https://github.com/python/black) [![Downloads](https://pepy.tech/badge/thinkcellbuilder/month)](https://pepy.tech/project/thinkcellbuilder) [![GitHub license](https://img.shields.io/github/license/philistino/thinkcellbuilder.svg)](https://github.com/philistino/thinkcellbuilder/blob/main/LICENSE) 

ThinkcellBuilder is a simple unofficial Python library used to build powerpoint presentations with Think-cell charts, textboxes, and tables. 

This project builds on [Duerto](https://github.com/duarteocarmo)'s [think-cell](https://github.com/duarteocarmo/think-cell) package. 

ThinkcellBuilder allows the user to automate data entry for all named Think-cell objects (e.g., charts, textfields, tables), except Gantt charts, on a powerpoint template. It also provides a presentation abstraction so that one can create a whole presentation using combinations of slide templates and charts.

This package outputs .ppttc files that, when opened with Think-cell, build powerpoint presentations. A [think-cell license and installation](https://www.think-cell.com/en/) license is not required to use this package, but one is needed to build the presentation. Pptc files are just 

### Installation

ThinkcellBuilder is available on PyPi. 

```sh
 $ pip install thinkcellbuilder
 ```

### Tutorial and usage

Let us say you have generated a template according to [think-cell's automation guidelines](https://www.think-cell.com/en/support/manual/jsondataautomation.shtml) called `simple-template.pptx` with the following chart called `Chart1`: 

<!-- <img src="https://raw.githubusercontent.com/duarteocarmo/think-cell/main/assets/example.png" width="500"> -->

The thinkcell library helps you generate a `.ppttc` file so that you can generate presentations based on that template using python:

```python
from thinkcellbuilder import Presentation, Template

# create a presentation object
presentation = Presentation()

# create template object, this usually represents one or a small number of slides
slide1 = Template("Company Performance Template.pptx")
# add your text field
slide1.add_textfield(
    name="Slide Title",
    text="Company Performance",
)
# add data for a chart named Chart1
slide1.add_chart(
    name="Chart1",
    categories=["Ads", "Revenue", "Losses"],
    data=[["Amazon", 1, 11, 14], ["Slack", 8, 2, 15], ["Uber", 1, 2, 12]],
)

# add slide1 to the presentation
presentation.add_template(slide1)

# create another slide from a different template
slide2 = Template("Company Forecast Template.pptx")
# add a text field
slide2.add_textfield(
    name="Slide Title",
    text="Tech Forecasts",
)
# add data for a chart named Chart1
slide2.add_chart(
    name="Chart1",
    categories=["3yr", "5yr", "10yr"],
    data=[["Amazon", 3, 10, 17], ["Slack", 8, 12, 15], ["Uber", 1, 2, 3]],
)

# add slide2 to the presentation
presentation.add_template(slide2)

# save the ppttc file 
presentation.save_ppttc(path="simple-example.ppttc")
 ```

Once done, go ahead and double click the generated `simple-example.ppttc` file. Think-cell will populate your charts/text fields based on the code above and open a powerpoint presentation. Save it and you are done!

You can also derive your chart from a Pandas dataframe. 

Make sure you have pandas installed (e.g., `pip install pandas`)

```python
from thinkcellbuilder import Presentation, Template
import pandas as pd

df = pd.DataFrame(
    columns=["Company", "Ads", "Revenue", "Losses"],
    data=[["Amazon", 1, 11, 14], ["Slack", 8, 2, 15], ["Ford", 1, 2, 12]],
)

slide = Template("simple-template.pptx") # create template object
slide.add_chart_from_dataframe(
    name="Chart1",
    dataframe=df,
) # add your dataframe

presentation = Presentation()
presentation.add_template(slide)
tc.save_ppttc(path="simple-example.ppttc")
 ```

Visit the [examples folder](examples) for more examples and source files. 

If you wish to learn more about this process, visit the think-cell [automation documentation](https://www.think-cell.com/en/support/manual/jsondataautomation.shtml). 

## Dependencies
ThinkcellBuilder has no dependencies outside the Python standard library. If you want to create charts from pandas dataframes, pandas is obviously needed.

## Limitations
It is currently impossible to derive the types or names of Think-cell objects (e.g., charts, named textfields, etc.) in a template programatically, so all automation methods rely on typo-free expression of object names. Mis-typed chart names are silently ignored by Think-cell when it is building the presentation, so please double check your object names and types. 

## Contributing

Start by forking this repo.


Install the development dependencies (you probably want to do this in a [virtual environment](https://docs.python-guide.org/dev/virtualenvs/)):

```shell
 $ pip install -r requirements-dev.txt
 ```

Make sure the tests run:

```shell
 $ pytest
 ```

Create a branch and submit a pull request. 


*Note: This project is in no way affiliated with think-cell Sales GmbH & Co. KG. I just wanted to make my (and hopefully your) life easier.*

