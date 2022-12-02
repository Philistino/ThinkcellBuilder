from __future__ import annotations

import json
import warnings
from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path
from typing import Optional, Union

try:
    import pandas as pd

    PANDAS = True
except ImportError:
    PANDAS = False
    warnings.warn(
        """Pandas not installed, methods adding charts from dataframes will raise RuntimeErrors.
    For full functionality install with 
        pip install thinkcell-builder[pandas]
    or make sure pandas is installed in the same environment
        'pip install pandas'"""
    )


class DataFrameError(Exception):
    pass


@dataclass
class Template:
    """Represents a template to be included in an output presentation.
    The .pptx file can include more than one slide.
    Charts can be added with the class methods if you need
    to populate specific parts of the template.

    It is not possible to verify:
        - that a template exists
        - that a template has named objects to populate
        - that objects are of the name or type described in these method calls

    Attributes
    ----------
    template : str
        the path to the template file. This can point to a url if
        your template can be downloaded from the internet. This includes
        network drives, dropbox links, etc.
    thinkcell_objects :  list[dict]
        This attribute is not available in the init function.
        Add objects via methods (e.g., add_textfield, add_table).
        a list of objects to update in the template
        each object will end up being an array of dictionaries like:
        "data": {
            "name": "chart1",
            "table": ["data to go in table"]
        }
    """

    path: str
    thinkcell_objects: list[dict] = field(init=False, default_factory=list)

    def add_textfield(self, name: str, text: str):
        """Adds a text field to the Template object.

        Parameters
        ----------
        name : str
            The name of the text field in the specified template
        text : str
            A string containing the text

        MIT License

        Copyright (c) 2019 Duarte OC
        """
        if not isinstance(name, str):
            warnings.warn(
                f"Your field name is not a string, we will convert it into one. But wanted to make sure you were aware.",
                UserWarning,
            )
        spec = {}
        spec["name"] = str(name)
        field_text = [self.transform_input(text)]
        spec["table"] = [field_text]
        self.thinkcell_objects.append(spec)

    def add_table(
        self, name: str, data: list[list], fill: Optional[list[str]] = None
    ):  # TODO make fills for each cell instead of each row, if fills really are possible
        """
        Adds a table to the template object.

        Parameters
        ----------
        name : str
            The name of the chart in the specified template
        data : list[list]
            A list of lists. Each list contains the row of data to be added.
        """
        if fill is None:
            fill = [None for _ in data]
        spec = {}
        spec["name"] = str(name)
        spec["table"] = []
        for data_list, color in zip(data, fill):
            spec["table"].append([self.transform_input(el, color) for el in data_list])
        self.thinkcell_objects.append(spec)

    @staticmethod
    def transform_input(data_element, color=None):
        """Transforms a `data element` into an object like {"type": data element}.

        Parameters
        ----------
        data_element : str, int, float, datetime
            A data element can be a string, int, float or datetime.
        color : str
            The hex or rgb string for the element color. If None, then
            no fill will be included.

        Returns
        -------
        dict
            Returns an object of type {"type": input}

        Raises
        ------
        ValueError
            Raises if object is not of type int, float, str, or datetime.

        Examples
        --------
        For a float or int, the dict will have "number" as key.

        >>> print(Template.transform_input(5))
        {"number": 5}

        For a str input, the dict will have a "string" as key.

        >>> print(Template.transform_input("test"))
        {"string": "test"}

        MIT License

        Copyright (c) 2019 Duarte OC
        """
        fill = {}
        if color is not None:
            fill = {"fill": color}

        if isinstance(data_element, datetime):
            return {"date": data_element.strftime("%Y-%m-%d"), **fill}

        if isinstance(data_element, str):
            return {"string": data_element, **fill}

        if isinstance(data_element, (int, float)):
            return {"number": data_element, **fill}
        else:
            raise ValueError(
                f"{data_element} of type {type(data_element)} is not acceptable."
            )

    def add_chart(
        self,
        name: str,
        categories: list[Union[str, int, float]],
        data: list[Union[str, int, float]],
        fill: Optional[list[str]] = None,
        first_row_blank: bool = True,
    ):
        """Adds a chart to the template object.
        This method works for the stacked, 100%, cluster, area, line, combination, Mekko charts

        Parameters
        ----------
        name : str
            The name of the chart in the specified template
        categories : list
            A list containing the header of the chart. Headers can
            be categories, years, companies, etc.
        data : list
            A list of lists. Each list contains the row of data to be added. Be
            aware that the first element of each of these lists should be a
            category as well.
        fill : list
            A list containing strings of either the hex or rgb values for fill
            for each series. Must match the length of the series. Can specify None
            to use no fill.
        first_row_blank : bool
            whether to add a blank row below the category labels in the chart's excel sheet.
            this row is mostly used for specifying a different number for 100% in stacked bar charts

        Raises
        ------
        ValueError
            If template does not exist, if the length of the categories does not
            make sense reltively to the header data.

        MIT License

        Copyright (c) 2019 Duarte OC
        """
        if not isinstance(name, str):
            warnings.warn(
                f"Your chart name is not a string, we will convert it into one. But wanted to make sure you were aware.",
                UserWarning,
            )

        for data_list in data:
            if len(data_list) != len(categories) + 1:
                raise ValueError(
                    f"Your categories should be the equal to the length of your data lists - 1. Your data element {data_list} is of size {len(data_list)} but should be of size {len(categories) + 1}."
                )
        if fill is not None and len(fill) != len(data):
            raise ValueError(
                f"Your fill colors should be the equal to the length of your data (the number of series). Your fill element {fill} is of size {len(fill)} but should be of size {len(data)}."
            )

        if fill is None:
            fill = [None for _ in data]

        spec = {}
        spec["name"] = str(name)
        chart_categories = [None] + [
            self.transform_input(element) for element in categories
        ]

        spec["table"] = [chart_categories]
        if first_row_blank:  # add blank row for 100% row
            spec["table"].append([])
        for data_list, color in zip(data, fill):
            spec["table"].append([self.transform_input(el, color) for el in data_list])
        self.thinkcell_objects.append(spec)

    def add_chart_from_dataframe(
        self,
        name: str,
        dataframe: pd.DataFrame,
        fill: Optional[list[str]] = None,
        first_row_blank: bool = True,
    ):
        """Adds a chart based on a dataframe to the template object.

        Parameters
        ----------
        name : str
            The name of the chart in the specified template
        dataframe : pandas.DataFrame
            The dataframe
        fill : list
            A list of strings the length of the number of series for specifying
            the fill colors with the hex or rgb
        first_row_blank : bool
            whether to add a blank row below the category labels in the chart's excel sheet.
            this row is mostly used for specifying a different number for 100% in stacked bar charts

        Raises
        ------
        DataFrameError
            If an invalid or empty DataFrame is passed

        MIT License

        Copyright (c) 2019 Duarte OC
        """
        if not PANDAS:
            raise ImportError(
                "You must have pandas installed to use this method. Install with 'pip install pandas'"
            )
        try:
            categories = dataframe.columns.to_list()[1:]
            assert isinstance(categories, list)
            data = dataframe.values.tolist()
            assert isinstance(data, list)
        except (AttributeError, AssertionError):
            raise DataFrameError("You did not pass a valid Pandas DataFrame")

        try:
            assert len(categories) >= 1
            assert len(data)
        except AssertionError:
            raise DataFrameError("The DataFrame you passed does not contain data")

        self.add_chart(
            name, categories, data, fill=fill, first_row_blank=first_row_blank
        )

    def add_pie_chart(
        self, name: str, data: list[list], fill: Optional[list[str]] = None
    ):
        """
        Adds a pie/doughnut chart to the Template object.
        The data parameter should be a list of two-element list(s)
        (e.g., [["label1", 0.1], ["label2", 0.9]])
        This can also stand in for Harvey balls if you only include
        one row and a value greater than 0.

        Parameters
        ----------
        name : str
            The name of the chart in the specified template
        data : list[list]
            A list of lists. Each list contains the row of data to be added.

        """
        for row in data:
            if len(row) != 2:
                raise RuntimeError(
                    f"Data not in correct shape for pie/doughnut chart {name}: {data}. Each row should contain two elements, one label and one number."
                )
        if fill is None:
            fill = [None for _ in data]
        spec = {}
        spec["name"] = str(name)
        spec["table"] = [[]]
        for data_list, color in zip(data, fill):
            spec["table"].append([self.transform_input(el, color) for el in data_list])
        self.thinkcell_objects.append(spec)

    def add_scatter_from_dataframe(
        self,
        name: str,
        dataframe,
        x: str,
        y: str,
        label: Optional[str] = None,
        size: Optional[str] = None,
        group: Optional[str] = None,
        fill: Optional[list[str]] = None,
    ):
        """Adds a scatter plot based on a dataframe to the template object.

        Parameters
        ----------
        name : str
            The name of the chart in the specified template
        dataframe : pandas.DataFrame
            The dataframe. Can include columns other than those used in this chart.
        x : str
            The column name for the x axis
        y : str
            The column name for the y axis
        label : Optional[str]
            The column name for the marker labels
        size : Optional[str]
            The column name for the one indicating the relative size of each marker
        group : Optional[str]
            The column name for the one indicating the style of marker
        fill : list
            A list of strings the length of the number of series for specifying
            the fill colors with the hex or rgb

        Raises
        ------
        DataFrameError
            If an invalid or empty DataFrame is passed
        """
        if not PANDAS:
            raise ImportError(
                "You must have pandas installed to use this method. Install with 'pip install pandas'"
            )
        df = dataframe.copy()
        for col in (label, size, group):
            if col is None:
                df[str(col)] = ""
        df = df[[str(label), x, y, str(size), str(group)]]
        self.add_chart_from_dataframe(name, df, fill, first_row_blank=False)

    def serialize(self) -> dict:
        """Returns dictionary representation of"""
        return {"template": self.path, "data": self.thinkcell_objects}


@dataclass
class Presentation:
    """Represents a whole presentation including one or more slides.
    Needs a template and charts.

    Add Template objects to the presentation object using the add_slide method
    and call the save_ppttc method when you are ready to save the ppttc file.

    Attributes
    ----------
    slides : list[Template]
        List of Template objects to concatenate in the ppttc file.
        This is an optional parameter in the init function.
        If the save_ppttc method is called without any slides,
    """

    slides: list[Template] = field(init=True, default_factory=list, repr=True)

    def _verify_template(self, template_name: str):
        """Verifies the validity of a template path.
        This does not verify if the .pptx file is accessible or if the .pptx file
        contains Think-cell objects

        Parameters
        ----------
        template_name : str
            The name of the template file to be added.

        Returns
        -------
        template_name: str
            Returns the name of the template if exceptions are not raised.

        Raises
        ------
        TypeError
            Raises if template name is not a string or does not end in ".pptx".

        MIT License

        Copyright (c) 2019 Duarte OC
        """
        if not isinstance(template_name, str):
            raise TypeError(f"'{template_name}' is not a valid template file.")

        if not template_name.endswith(".pptx"):
            raise TypeError(f"'{template_name}' is not a valid Powerpoint file.")

        else:
            return template_name

    def add_template(self, template: Template):
        """
        Add a populated template to the presetation

        Parameters
        ----------
        template : Template
            The template representing slides to be added to the presentation.

        MIT License

        Copyright (c) 2019 Duarte OC
        """
        self._verify_template(template.path)
        self.slides.append(template)

    def save_ppttc(self, filename: Union[Path, str]):
        """Saves the Thinkcell object as a `.ppttc` file.

        Parameters
        ----------
        filename : str
            The name of the file to be saved.

        Raises
        ------
        ValueError
            If the filename specified is not a string or does
            not end in `.ppttc`.
        ValueError
            If the presentation has no slides

        MIT License

        Copyright (c) 2019 Duarte OC
        """
        if not str(filename).endswith(".ppttc"):
            raise ValueError(
                f"You want to save your file as a '.ppttc' file, not a '{filename}'. Visit https://www.think-cell.com/en/support/manual/jsondataautomation.shtml for more information."
            )

        if not self.slides:
            raise ValueError(
                f"Please add data before saving to a template file by using 'add_slide'."
            )

        to_dump = [i.serialize() for i in self.slides]
        with open(filename, "w") as outfile:
            json.dump(to_dump, outfile)
            return True
