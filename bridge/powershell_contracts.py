"""Shared PowerShell fragments that implement public Action Contract enums."""


CHART_TYPE_VALUES = {
    "column": 51,
    "bar": 57,
    "line": 4,
    "pie": 5,
    "doughnut": -4120,
    "area": 1,
    "scatter": -4169,
}

CHART_TYPE_ACTIONS = {
    "excel": "createChart",
    "ppt": "insertPptChart",
}


def render_chart_type_converter():
    cases = "\n".join(
        f'            "{name}" {{ return {value} }}'
        for name, value in CHART_TYPE_VALUES.items()
    )
    return f'''function Convert-ChartType($value) {{
    if ($value -is [string]) {{
        switch ($value.ToLowerInvariant()) {{
{cases}
            default {{ throw "未知 chartType: $value" }}
        }}
    }}
    return [int]$value
}}
'''
