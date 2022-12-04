import os
from datetime import datetime

import pandas as pd
import pytest

from thinkcellbuilder import DataFrameError, Presentation, Template


def test_init():
    tc = Template("template.pptx")
    assert tc.thinkcell_objects == []


@pytest.mark.parametrize(
    "test_input, expected",
    [
        ("daf", {"string": "daf"}),
        (3, {"number": 3}),
        (2.0, {"number": 2.0}),
        (datetime(2012, 9, 16, 0, 0), {"date": "2012-09-16"}),
    ],
)
def test_transform_input(test_input, expected):
    assert Template.transform_input(test_input) == expected


def test_transform_input_bad():
    with pytest.raises(ValueError):
        Template.transform_input([3, 4])


def test_verify_template_1():
    template_name = "not a file name"
    with pytest.raises(TypeError):
        Presentation()._verify_template(template_name)


def test_verify_template_2():
    template_name = 5
    with pytest.raises(TypeError):
        Presentation()._verify_template(template_name)


def test_verify_template_3():
    template_name = "example.pptx"
    assert Presentation()._verify_template(template_name) == template_name


def test_add_chart_warning():
    tc = Template(path="template.pptx")
    with pytest.warns(UserWarning) as record:
        tc.add_chart(
            name=234,
            categories=["Alpha", "bravo"],
            data=[[3, 4, datetime(2012, 9, 16, 0, 0)], [2, "adokf", 6]],
        )


def test_add_textfield_warning():
    tc = Template(path="template.pptx")
    with pytest.warns(UserWarning) as record:
        tc.add_textfield(
            name=234,
            text="A great slide",
        )


def test_add_chart_bad_dimensions():
    tc = Template(path="example.pptx")
    with pytest.raises(ValueError):
        tc.add_chart(
            name="Cool Name bro",
            categories=["Alpha", "bravo"],
            data=[[3, 4, datetime(2012, 9, 16, 0, 0)], [2, "adokf"]],
        )


def test_add_chart():
    tc = Template(path="example.pptx")
    tc.add_chart(
        name="Cool Name bro",
        categories=["Alpha", "bravo"],
        data=[[3, 4, datetime(2012, 9, 16, 0, 0)], [2, "adokf", 4]],
    )
    got = tc.serialize()
    want = {
        "template": "example.pptx",
        "data": [
            {
                "name": "Cool Name bro",
                "table": [
                    [None, {"string": "Alpha"}, {"string": "bravo"}],
                    [],
                    [
                        {"number": 3},
                        {"number": 4},
                        {"date": "2012-09-16"},
                    ],
                    [
                        {"number": 2},
                        {"string": "adokf"},
                        {"number": 4},
                    ],
                ],
            },
        ],
    }
    assert got == want


def test_add_chart_with_fill():
    tc = Template(path="example.pptx")
    tc.add_chart(
        name="Cool Name bro",
        categories=["Alpha", "bravo"],
        data=[[3, 4, datetime(2012, 9, 16, 0, 0)], [2, "adokf", 4]],
        fill=["#70AD47", "#ED7D31"],
    )

    assert tc.serialize() == {
        "template": "example.pptx",
        "data": [
            {
                "name": "Cool Name bro",
                "table": [
                    [None, {"string": "Alpha"}, {"string": "bravo"}],
                    [],
                    [
                        {"number": 3, "fill": "#70AD47"},
                        {"number": 4, "fill": "#70AD47"},
                        {"date": "2012-09-16", "fill": "#70AD47"},
                    ],
                    [
                        {"number": 2, "fill": "#ED7D31"},
                        {"string": "adokf", "fill": "#ED7D31"},
                        {"number": 4, "fill": "#ED7D31"},
                    ],
                ],
            }
        ],
    }


def test_add_chart_with_fill_error():
    tc = Template(path="example.pptx")
    with pytest.raises(ValueError):
        tc.add_chart(
            name="Cool Name bro",
            categories=["Alpha", "bravo"],
            data=[[3, 4, datetime(2012, 9, 16, 0, 0)], [2, "adokf", 4]],
            fill=["#70AD47"],
        )


def test_add_chart_from_dataframe():
    tc = Template(path="example.pptx")
    dataframe = pd.DataFrame(
        columns=["Company", "Employees", "Revenue", "Other"],
        data=[
            ["Apple", 200, 1.5, 10],
            ["Amazon", 100, 1.0, 12],
            ["Slack", 50, 0.5, 16],
        ],
    )
    tc.add_chart_from_dataframe(
        name="Cool Chart",
        dataframe=dataframe,
    )
    assert tc.serialize() == {
        "template": "example.pptx",
        "data": [
            {
                "name": "Cool Chart",
                "table": [
                    [
                        None,
                        {"string": "Employees"},
                        {"string": "Revenue"},
                        {"string": "Other"},
                    ],
                    [],
                    [
                        {"string": "Apple"},
                        {"number": 200},
                        {"number": 1.5},
                        {"number": 10},
                    ],
                    [
                        {"string": "Amazon"},
                        {"number": 100},
                        {"number": 1.0},
                        {"number": 12},
                    ],
                    [
                        {"string": "Slack"},
                        {"number": 50},
                        {"number": 0.5},
                        {"number": 16},
                    ],
                ],
            }
        ],
    }


def test_add_chart_from_dataframe_with_fill():
    tc = Template(path="example.pptx")
    dataframe = pd.DataFrame(
        columns=["Company", "Employees", "Revenue", "Other"],
        data=[
            ["Apple", 200, 1.5, 10],
            ["Amazon", 100, 1.0, 12],
            ["Slack", 50, 0.5, 16],
        ],
    )
    tc.add_chart_from_dataframe(
        name="Cool Chart",
        dataframe=dataframe,
        fill=["#70AD47", "#ED7D31", "#4472C4"],
    )
    assert tc.serialize() == {
        "template": "example.pptx",
        "data": [
            {
                "name": "Cool Chart",
                "table": [
                    [
                        None,
                        {"string": "Employees"},
                        {"string": "Revenue"},
                        {"string": "Other"},
                    ],
                    [],
                    [
                        {"string": "Apple", "fill": "#70AD47"},
                        {"number": 200, "fill": "#70AD47"},
                        {"number": 1.5, "fill": "#70AD47"},
                        {"number": 10, "fill": "#70AD47"},
                    ],
                    [
                        {"string": "Amazon", "fill": "#ED7D31"},
                        {"number": 100, "fill": "#ED7D31"},
                        {"number": 1.0, "fill": "#ED7D31"},
                        {"number": 12, "fill": "#ED7D31"},
                    ],
                    [
                        {"string": "Slack", "fill": "#4472C4"},
                        {"number": 50, "fill": "#4472C4"},
                        {"number": 0.5, "fill": "#4472C4"},
                        {"number": 16, "fill": "#4472C4"},
                    ],
                ],
            }
        ],
    }


def test_add_chart_from_dataframe_invalid_dataframe():
    tc = Template(path="example.pptx")
    dataframe = [
        ["Apple", 200, 1.5, 10],
        ["Amazon", 100, 1.0, 12],
        ["Slack", 50, 0.5, 16],
    ]
    with pytest.raises(DataFrameError):
        tc.add_chart_from_dataframe(
            name="Cool Chart",
            dataframe=dataframe,
        )


def test_add_chart_from_dataframe_no_columns():
    tc = Template(path="example.pptx")
    dataframe = pd.Series(
        data=[
            ["Apple", 200, 1.5, 10],
            ["Amazon", 100, 1.0, 12],
            ["Slack", 50, 0.5, 16],
        ]
    )
    with pytest.raises(DataFrameError):
        tc.add_chart_from_dataframe(
            name="Cool Chart",
            dataframe=dataframe,
        )


def test_add_chart_from_dataframe_no_data():
    tc = Template(path="example.pptx")
    dataframe = pd.DataFrame(
        columns=["Company"], data=[["Apple"], ["Amazon"], ["Slack"]]
    )
    with pytest.raises(DataFrameError):
        tc.add_chart_from_dataframe(
            name="Cool Chart",
            dataframe=dataframe,
        )


def test_add_chart_from_dataframe_no_rows():
    tc = Template(path="example.pptx")
    dataframe = pd.DataFrame(columns=["Company", "Employees", "Revenue", "Other"])
    with pytest.raises(DataFrameError):
        tc.add_chart_from_dataframe(
            name="Cool Chart",
            dataframe=dataframe,
        )


def test_add_textfield():
    tc = Template(path="example.pptx")
    tc.add_textfield(name="Title", text="A great slide")
    assert tc.serialize() == {
        "template": "example.pptx",
        "data": [{"name": "Title", "table": [[{"string": "A great slide"}]]}],
    }


@pytest.mark.parametrize("input, output", [("word.docx", ValueError), (3, ValueError)])
def test_save_ppttc_bad_file(input, output):
    tc = Template(path="example.pptx")
    tc.add_chart(
        name="Chart name",
        categories=["alpha", "bravo"],
        data=[["today", 1, 2], ["tomorrow", 3, 4]],
    )
    with pytest.raises(output):
        Presentation([tc]).save_ppttc(path=input)


def test_save_pptc():
    tc = Presentation()
    with pytest.raises(ValueError):
        tc.save_ppttc("test.ppttc")


def test_save_ppttc():
    tc = Template(path="example.pptx")
    tc.add_chart(
        name="Chart name",
        categories=["alpha", "bravo"],
        data=[["today", 1, 2], ["tomorrow", 3, 4]],
    )
    assert Presentation([tc]).save_ppttc(path="test.ppttc") == True
    os.remove("test.ppttc")


def test_add_table():
    tc = Template(path="example.pptx")
    tc.add_table(
        name="nice table",
        data=[["A1", "A2"], ["B1", "B2"]],
        fill=["#000000", "#111111"],
    )
    assert tc.serialize() == {
        "template": "example.pptx",
        "data": [
            {
                "name": "nice table",
                "table": [
                    [
                        {"string": "A1", "fill": "#000000"},
                        {"string": "A2", "fill": "#000000"},
                    ],
                    [
                        {"string": "B1", "fill": "#111111"},
                        {"string": "B2", "fill": "#111111"},
                    ],
                ],
            }
        ],
    }


# def test_add_table_per_cell_fill():
#     """come back to this if table fills are possible"""
#     tc = Template(path="example.pptx")
#     tc.add_table(
#         name="nice table",
#         data=[["A1", "A2"], ["B1", "B2"]],
#         fill=[
#             ["#000000", "#111111"],
#             ["#222222", "#333333"],
#         ],
#     )
#     assert tc.serialize() == {
#         "template": "example.pptx",
#         "data": [
#             {
#                 "name": "nice table",
#                 "table": [
#                     [
#                         {"string": "A1", "fill": "#000000"},
#                         {"string": "A2", "fill": "#111111"},
#                     ],
#                     [
#                         {"string": "B1", "fill": "#222222"},
#                         {"string": "B2", "fill": "#333333"},
#                     ],
#                 ],
#             }
#         ],
#     }


def test_add_table_nofill():
    tc = Template(path="example.pptx")
    tc.add_table(
        name="nice table",
        data=[["A1", "A2"], ["B1", "B2"]],
    )
    assert tc.serialize() == {
        "template": "example.pptx",
        "data": [
            {
                "name": "nice table",
                "table": [
                    [
                        {"string": "A1"},
                        {"string": "A2"},
                    ],
                    [
                        {"string": "B1"},
                        {"string": "B2"},
                    ],
                ],
            }
        ],
    }


def test_add_pie_chart_nofill():
    tc = Template(path="example.pptx")
    tc.add_pie_chart(
        name="tasty pie",
        data=[["Series1", 90], ["Series2", 10]],
    )
    assert tc.serialize() == {
        "template": "example.pptx",
        "data": [
            {
                "name": "tasty pie",
                "table": [
                    [],
                    [{"string": "Series1"}, {"number": 90}],
                    [{"string": "Series2"}, {"number": 10}],
                ],
            }
        ],
    }


def test_add_pie_chart_fill():
    tc = Template(path="example.pptx")
    tc.add_pie_chart(
        name="tasty pie",
        data=[["Series1", 90], ["Series2", 10]],
        fill=["#000000", "#111111"],
    )
    assert tc.serialize() == {
        "template": "example.pptx",
        "data": [
            {
                "name": "tasty pie",
                "table": [
                    [],
                    [
                        {"string": "Series1", "fill": "#000000"},
                        {"number": 90, "fill": "#000000"},
                    ],
                    [
                        {"string": "Series2", "fill": "#111111"},
                        {"number": 10, "fill": "#111111"},
                    ],
                ],
            }
        ],
    }


def test_add_pie_chart_error():
    tc = Template(path="example.pptx")
    with pytest.raises(RuntimeError):
        tc.add_pie_chart(name="tasty pie", data=[["Series1", 90], ["Series2"]])


def test_add_slide():
    tc = Template(path="example.pptx")
    tc.add_pie_chart(name="tasty pie", data=[["Series1", 90], ["Series2", 10]])
    pres = Presentation()
    pres.add_template(tc)
    assert pres.slides[0] == tc


def test_add_scatter():
    slide = Template(path="example.pptx")
    dataframe = pd.DataFrame(
        columns=["Company", "Employees", "Revenue", "Other"],
        data=[
            ["Apple", 200, 1.5, 10],
            ["Amazon", 100, 1.0, 12],
            ["Slack", 50, 0.5, 16],
        ],
    )
    slide.add_scatter_from_dataframe(
        name="Scatter Batter",
        dataframe=dataframe,
        x="Employees",
        y="Revenue",
        label="Company",
    )
    assert slide.serialize() == {
        "template": "example.pptx",
        "data": [
            {
                "name": "Scatter Batter",
                "table": [
                    [
                        None,
                        {"string": "Employees"},
                        {"string": "Revenue"},
                        {"string": "None"},
                        {"string": "None"},
                    ],
                    [
                        {"string": "Apple"},
                        {"number": 200},
                        {"number": 1.5},
                        {"string": ""},
                        {"string": ""},
                    ],
                    [
                        {"string": "Amazon"},
                        {"number": 100},
                        {"number": 1.0},
                        {"string": ""},
                        {"string": ""},
                    ],
                    [
                        {"string": "Slack"},
                        {"number": 50},
                        {"number": 0.5},
                        {"string": ""},
                        {"string": ""},
                    ],
                ],
            }
        ],
    }
