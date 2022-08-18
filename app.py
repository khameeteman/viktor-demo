"""
This file is the entry point for your application and is used to:

    - Define all entity-types that are part of the app, and
    - Create 1 or more initial entities (of above type(s)), which are generated upon starting the app

For more information about this file, see: https://docs.viktor.ai/docs/guides/fundamentals/app-file
"""
# Import required classes and functions
from pathlib import Path

from viktor import File, Color, ViktorController, UserException, InitialEntity
from viktor.external.word import WordFileTag, WordFileImage, render_word_file
from viktor.external.spreadsheet import SpreadsheetCalculation, SpreadsheetCalculationInput
from viktor.geometry import CircularExtrusion, SquareBeam, Extrusion, Material, Line, Point
from viktor.result import DownloadResult
from viktor.parametrization import ViktorParametrization, Step, GeoPointField, NumberField, DownloadButton, DateField, \
    IntegerField, OptionField, Table, TextField, TextAreaField, FileField, DynamicArray, OptionListElement, Text, \
    Section, BooleanField, Lookup
from viktor.utils import convert_word_to_pdf
from viktor.views import MapView, MapPoint, MapLegend, MapLabel, MapResult, PDFView, PDFResult, GeometryView, \
    GeometryResult, DataGroup, DataItem, DataStatus, PNGView, PNGResult, PlotlyAndDataView, PlotlyAndDataResult, Label


# Define some constants to be used in multiple parts of the code
NORM_A_MAX = 500
NORM_B_MAX = 750
NORM_C_MAX = 1000


# Define some stand-alone functions
def get_color(value: int) -> Color:
    """ Generate a color based on a single integer value (0 <= value <= 100)

    See "tests/test_get_color.py" for an example test written for this method.
    """
    if not 0 <= value <= 100:  # check for invalid value
        raise ValueError(f"value ({value}) must be between 0 - 100")

    return Color(int(value / 100 * 255), 255, 0)


def calculate_mass_from_spreadsheet(volume: float, density: float) -> float:
    """ Calculates the mass from the volume and density by means of a spreadsheet (calculate_mass.xlsx).

    For more information, see: https://docs.viktor.ai/docs/guides/services/spreadsheet-calculator
    """
    inputs = [
        SpreadsheetCalculationInput('volume', volume),  # 'volume' equals the name in viktor-input-sheet
        SpreadsheetCalculationInput('density', density),  # 'density' equals the name in viktor-input-sheet
    ]

    sheet_path = Path(__file__).parent / 'calculate_mass.xlsx'
    sheet = SpreadsheetCalculation.from_path(sheet_path, inputs=inputs)
    spreadsheet_result = sheet.evaluate(include_filled_file=False)
    return spreadsheet_result.get_value('mass')  # 'mass' equals the name in viktor-output-sheet


class Parametrization(ViktorParametrization):
    """ A Parametrization defines the fields and views visible in an entity type's editor.

    For more information on the Parametrization-class, see:
    https://docs.viktor.ai/docs/guides/fundamentals/parametrization-class

    Available fields are:

        - Text
        - TextField
        - TextAreaField
        - NumberField
        - IntegerField
        - DateField
        - BooleanField
        - OutputField
        - HiddenField
        - LineBreak
        - OptionField
        - MultiSelectField
        - AutocompleteField
        - ActionButton
        - DownloadButton
        - OptimizationButton
        - SetParamsButton
        - Table
        - DynamicArray
        - EntityOptionField
        - ChildEntityOptionField
        - SiblingEntityOptionField
        - EntityMultiSelectField
        - ChildEntityMultiSelectField
        - SiblingEntityMultiSelectField
        - GeoPointField
        - GeoPolylineField
        - GeoPolygonField
        - FileField
        - MultiFileField

    Structuring of fields can be achieved by means of Tab(s), Page(s) or Step(s) to create top level, and Section(s) to
    create a second level of layering.

    If the Parametrization-class becomes large, it might be cleaner to move the class to a separate file
    ("parametrization.py") next to "controller.py" and import the class in the usual way:

    from .parametrization import Parametrization

    For more information on an app's recommended folder structure, see:
    https://docs.viktor.ai/docs/guides/fundamentals/folder-structure#recommended-structure
    """
    research = Step("Research", views=['map_view'])
    research.data = Section("Data")
    research.data.text = Text(
        "Welcome in the editor. In general, an editor has input (parametrization) on the left and output (views) on "
        "the right. Some fields in the parametrization and views have an (i) next to it. Hover over it to get "
        "additional info.\\\n"
        "\\\n"
        "Play around with the value of 'Measurement' and 'Location', and see the results in the 'Map' view."
    )
    research.data.measurement = IntegerField(
        "Measurement", min=0, max=100,
        description="When removing the value with backspace, the value returned in the params is None."
    )
    research.data.location = GeoPointField(
        "Location",
        description="Remove the pre-selected point by clicking the bin icon. Define a new point by clicking the "
                    "+marker icon and selecting a location on the Map view."
    )

    research.section = Section("This is a section (click to expand/collapse)")
    research.section.text = Text(
        "A simple parametrization may consist of only fields. If the parametrization becomes large, a layered "
        "structure can be achieved by making use of Tab(s), "
        "[Page(s)](https://docs.viktor.ai/docs/guides/fundamentals/parametrization-class#creating-a-page) or "
        "[Step(s)](https://docs.viktor.ai/docs/guides/fundamentals/parametrization-class#creating-a-step) (used in "
        "this demo). A second level of layering can be achieved by using Section(s) (shown here).\\\n"
        "\\\n"
        "Click 'Next step' to continue.")

    design = Step("Design", views=['geometry_view'])
    design.text1 = Text(
        "There are many different "
        "[fields](https://docs.viktor.ai/docs/guides/fundamentals/parametrization-class#overview-of-input-fields) and "
        "[views](https://docs.viktor.ai/docs/guides/fundamentals/views#supported-views) available for the design of "
        "your editor."
    )
    design.shape = OptionField("Shape", options=['Circle', 'Rectangle', 'Triangle'], default='Circle', flex=50,
                               variant='radio')
    design.height = NumberField(
        "Height", suffix='m', min=1, max=10,
        description="min (1) and max (10) has been set on this field. Type in a value outside these limits and notice "
                    "that a warning is shown on the field and in the top bar.\\\n"
                    "\\\n"
                    "Empty to see a custom user-error raised."
    )
    design.red = NumberField("Red", min=0, max=255, step=1, default=100, variant='slider', flex=100)
    design.green = NumberField("Green", min=0, max=255, step=1, default=100, variant='slider', flex=100)
    design.blue = NumberField("Blue", min=0, max=255, step=1, default=100, variant='slider', flex=100)
    design.text2 = Text(
        "Many field settings can be "
        "[dynamic](https://docs.viktor.ai/docs/guides/topics/parametrization/set-field-constraints). For example, it "
        "is possible to set the visibility of a field based on the value of another field.")
    design.show_label = BooleanField("Show label", default=True, flex=15)
    design.label = TextField("Label", visible=Lookup('design.show_label'),  # only visible if show_label=True
                             description="This field is only visible if 'Show label' = True.")
    design.text3 = Text("This view has been set to update immediately when input is changed. This is especially useful "
                        "if calculating the results from the input doesn't take long.")

    calculate = Step("Calculate", views=['plotly_and_data_view'])
    calculate.text1 = Text(
        "Dynamically sized input can be realized using [Table(s) or DynamicArray(s)]"
        "(https://docs.viktor.ai/docs/guides/fundamentals/parametrization-class#table-vs-dynamicarray)."
    )
    calculate.cases = DynamicArray(
        "Cases", min=1, row_label="Case",
        description="Delete a row in a DynamicArray by clicking the bin icon. Add a row by clicking the + (add new "
                    "row) button.")
    calculate.cases.volume = NumberField("Volume", suffix='m³', min=0.1, max=1, step=0.1, default=0.3)
    calculate.cases.density = IntegerField("Density", suffix='kg/m³', min=0, max=3000, default=1000)
    calculate.cases.norm = OptionField(
        "Norm", options=[OptionListElement('A', f'A (max. {NORM_A_MAX} kg)'),
                         OptionListElement('B', f'B (max. {NORM_B_MAX} kg)'),
                         OptionListElement('C', f'C (max. {NORM_C_MAX} kg)')], default='A'
    )
    calculate.text2 = Text(
        "Besides plain Python code, calculations within a VIKTOR app can also be performed by making use of the "
        "[spreadsheet calculator service](https://docs.viktor.ai/docs/guides/services/spreadsheet-calculator). "
        "Furthermore, VIKTOR provides various [integrations](https://docs.viktor.ai/docs/guides/integrations/) with "
        "third-party software packages (note: integrations require a [worker](https://docs.viktor.ai/docs/worker), "
        "which is available to premium users only)."
    )
    calculate.spreadsheet = BooleanField(
        "Calculation method\\\n(Python function / Spreadsheet)", default=False,
        description="False = run calculation using simple Python function.\\\n"
                    "\\\n"
                    "True = run calculation using the spreadsheet calculator service."
    )
    calculate.text3 = Text(
        "This view has been set to update only when the user clicks the 'Update' button. This is especially useful if "
        "calculating the results from the input takes a long time (e.g. when using a service or integration that runs "
        "an external calculation tool). Click the 'Update' button on the 'Results' view to run the calculation."
    )

    report = Step("Report", views=['pdf_view'])
    report.download = Section("Download")
    report.download.text = Text(
        "There are various buttons available to let the user perform some action. For example, a "
        "[DownloadButton](https://docs.viktor.ai/docs/guides/topics/files#downloading) lets the user download a file, "
        "such as a report, to his/her local hard drive."
    )
    report.download.date = DateField("Date", description="Type or select a date by clicking on the calendar icon.")
    report.download.authors = Table("Authors", description="Right-click on a cell in the table to add or remove a row.")
    report.download.authors.first_name = TextField("First Name")
    report.download.authors.last_name = TextField("Last Name")
    report.download.authors.organization = TextField("Organization")
    report.download.authors.email = TextField("Email")
    report.download.remarks = TextAreaField("Additional Remarks", default="No remarks.")
    report.download.button = DownloadButton("Download Report (.docx)", method="download_report")
    report.upload = Section("Upload")
    report.upload.text = Text(
        "A [FileField](https://docs.viktor.ai/docs/guides/topics/files#uploading) enables the user to upload a file "
        "directly to the app.\\\n"
        "\\\n"
        "**Challenge**: copy \"report_template.docx\" from the demo app's source code, modify it to your liking, "
        "and upload it using this file-field, to generate and download a report that differs from the default one."
    )
    report.upload.template = FileField("Custom Template (Optional)", file_types=['.docx'], flex=50,
                                       description="Don't forget to select the file after uploading it.")

    evaluate = Step("Evaluate")
    evaluate.text = Text(
        "We hope you enjoyed this demo!\\\n"
        "\\\n"
        "Have a look add this demo's source code by opening it in your favorite "
        "[IDE](https://docs.viktor.ai/docs/guides/tools/ide).\\\n"
        "The code includes plenty of examples and comments to help you understand and is a good starting point for "
        "creating your own app.\\\n"
        "\\\n"
        "If you run into any questions that are not covered by the "
        "[documentation](https://docs.viktor.ai/docs/welcome), feel free to ask your questions on our "
        "[Community](https://community.viktor.ai/) page."
    )

    ########
    # Uncomment the lines below to add a step that includes a text-field (and has no views)
    ########
    # my_step = Step("My Step")
    # my_step.text = Text("From the moment you uncommented the 'Text' field, it became visible in the editor!")


# Creates an entity-type 'MyEntityType'
class MyEntityType(ViktorController):
    """
    For more information on the Controller-class, see:
    https://docs.viktor.ai/docs/guides/fundamentals/controller-class

    Views are defined by means of "@***View" decorated methods on the Controller.

    Available views are:

        - GeometryView (2D/3D)
        - DataView
        - SVGView
        - PNGView
        - JPGView
        - MapView
        - GeoJSONView
        - WebView
        - PlotlyView
        - PDFView
        - GeometryAndDataView
        - SVGAndDataView
        - PNGAndDataView
        - JPGAndDataView
        - MapAndDataView
        - GeoJSONAndDataView
        - WebAndDataView
        - PlotlyAndDataView

    For more information on views, see: https://docs.viktor.ai/docs/guides/fundamentals/views
    """
    label = 'My Entity Type'  # label of the entity type as seen by the user in the app's interface
    parametrization = Parametrization  # parametrization associated with the editor of the MyEntityType entity type

    @MapView("Map", duration_guess=1, description="Click on the map marker to see its description.")
    def map_view(self, params, **kwargs):
        """ https://docs.viktor.ai/docs/guides/topics/views/create-map-view """
        data = params.research.data

        location = data.location  # type GeoPoint (or None if unfilled)

        features, labels = [], []
        if location:
            measurement = data.measurement
            if measurement is not None:  # colored pin-add icon if measurement exists
                color = get_color(measurement)
                icon = "pin-add"

                label = MapLabel(location.lat, location.lon, f"{measurement}", scale=17)
                labels.append(label)
            else:  # black cross icon if no measurement
                color = Color.black()  # black map marker if no measurement
                icon = "cross"  # cross icon if no measurement

            description = f"Measurement: {measurement}"
            map_point = MapPoint.from_geo_point(location, description=description, color=color, icon=icon)
            features.append(map_point)

        legend_entries = [(get_color(i), f"{i}") for i in [0, 10, 20, 30, 40, 50, 60, 70, 80, 90, 100]]
        legend = MapLegend(legend_entries)

        return MapResult(features, labels, legend)

    @GeometryView("Geometry", duration_guess=1, description="Move around and zoom in and out using your mouse.")
    def geometry_view(self, params, **kwargs):
        """ https://docs.viktor.ai/docs/guides/topics/views/create-geometry-view """
        design_params = params.design

        # Create material with the user-defined color
        red = design_params.red
        green = design_params.green
        blue = design_params.blue
        material = Material("my_material", color=Color(red, green, blue))

        shape = design_params.shape
        height = design_params.height
        if height is not None:  # height is filled in
            if shape == 'Circle':
                extrusion_line = Line(Point(0, 0, -height / 2), Point(0, 0, height / 2))
                geometry = CircularExtrusion(diameter=1, line=extrusion_line, material=material)
            elif shape == 'Rectangle':
                geometry = SquareBeam(1, 1, height, material=material)
            elif shape == 'Triangle':
                profile = [Point(1, 0), Point(-1, 0), Point(0, 2), Point(1, 0)]
                extrusion_line = Line(Point(0, 0, -height / 2), Point(0, 0, height / 2))
                geometry = Extrusion(profile, extrusion_line, material=material)
            else:
                raise NotImplementedError  # should not occur
        else:  # raise an exception to the user if height is not filled in
            raise UserException("Please fill in a value for 'height'")

        labels = []
        if design_params.show_label and design_params.label:
            label = Label(Point(0, 0, -height), design_params.label, size_factor=2)
            labels.append(label)

        return GeometryResult(geometry, labels=labels)

    @PlotlyAndDataView("Results", duration_guess=4,  # adds 'Update' button if duration_guess > 3
                       description="Most of the views can be combined with a dataview (e.g. PlotlyView + DataView = "
                                   "PlotlyAndDataView), to have a more compact overview of the results.")
    def plotly_and_data_view(self, params, **kwargs):
        """ Combined PlotlyView + DataView.

        More information on PlotlyView: https://docs.viktor.ai/docs/guides/topics/views/create-plotly-view
        More information on DataView: https://docs.viktor.ai/docs/guides/topics/views/create-data-view
        """
        graph_data, data_items = [], []
        for i, case in enumerate(params.calculate.cases, 1):
            if case.norm == 'A':
                max_mass = NORM_A_MAX
            elif case.norm == 'B':
                max_mass = NORM_B_MAX
            elif case.norm == 'C':
                max_mass = NORM_C_MAX
            else:
                raise NotImplementedError  # should not occur

            if params.calculate.spreadsheet:
                # calculate the mass using a spreadsheet
                mass = calculate_mass_from_spreadsheet(case.volume, case.density)
            else:
                # calculate the mass using a simple Python function (very fast)
                mass = case.volume * case.density

            unity_check = mass / max_mass * 100

            if unity_check > 100:
                status = DataStatus.ERROR
                color = "red"
            elif unity_check > 80:
                status = DataStatus.WARNING
                color = "orange"
            else:
                status = DataStatus.SUCCESS
                color = "green"

            graph_data.append((unity_check, color))

            item = DataItem(
                f"Case {i}", unity_check, suffix='%', number_of_decimals=0, status=status,
                subgroup=DataGroup(
                    DataItem("Volume", case.volume, suffix='m³'),
                    DataItem("Density", case.density, suffix='kg/m³'),
                    DataItem("Mass", mass, suffix='kg'),
                    DataItem("Norm", case.norm),
                    DataItem("Unity Check", unity_check, suffix='%', number_of_decimals=0, status=status)
                )
            )

            data_items.append(item)

        # Create the Plotly graph
        x = [f"Case {i}" for i in range(1, len(graph_data) + 1)]    # x = Case 1, Case 2, Case 3, ...
        if graph_data:
            y, color = zip(*graph_data)                             # y = unity checks
        else:
            raise UserException("Add at least 1 case.")
        # For more information on how to create a Plotly visualization, see:
        # https://plotly.com/chart-studio-help/json-chart-schema/
        fig = {
            'data': [{'type': "bar", 'x': x, 'y': y, 'marker': {'color': color}}],
            'layout': {'title': {'text': "Unity Check [%]"}}
        }

        return PlotlyAndDataResult(fig, data=DataGroup(*data_items))

    @staticmethod
    def create_report(params, entity_name: str) -> File:
        """ Create a report using a Word-file template (report_template.docx).

        For more information, see: https://docs.viktor.ai/docs/guides/services/word-file-templater
        """
        content = params.report.download

        # DateField returns a Python datetime object
        date = content.date.strftime("%Y-%m-%d") if content.date is not None else '-'

        components = [
            WordFileTag('title', entity_name),
            WordFileTag('date', date),
            WordFileTag('authors', content.authors),
            WordFileTag('remarks', content.remarks)
        ]

        with open(Path(__file__).parent / 'viktor-logo.png', 'rb') as image:
            components.append(WordFileImage(image, 'image', width=30))

        template = params.report.upload.template
        if template is not None:  # user has selected custom report template
            with template.file.open_binary() as r:
                report = render_word_file(r, components)
        else:  # user has not selected custom report template, use default template
            with open(Path(__file__).parent / 'report_template.docx', 'rb') as r:
                report = render_word_file(r, components)

        return report

    def download_report(self, params, entity_name: str, **kwargs):
        """ Enables the user to download a report.

        For more information on downloading and uploading files, see: https://docs.viktor.ai/docs/guides/topics/files
        """
        report = self.create_report(params, entity_name)
        return DownloadResult(report, "report.docx")

    @PDFView("PDF", duration_guess=10, update_label="Generate Report",
             description="The PDFView makes it possible to show a static or dynamically generated report directly in "
                         "your VIKTOR app.")
    def pdf_view(self, params, entity_name: str, **kwargs):
        """ https://docs.viktor.ai/docs/guides/topics/views/create-pdf-view """
        docx_report = self.create_report(params, entity_name)

        with docx_report.open_binary() as r:
            pdf_report = convert_word_to_pdf(r)

        return PDFResult(file=pdf_report)

    ########
    # Uncomment the lines below to add a PNG (image) view. Add the method name ('png_view') to the 'views' argument in
    # any of the existing steps in the Parametrization to make the view visible in that step.
    ########
    # @PNGView("Image", duration_guess=1)
    # def png_view(self, params, **kwargs):
    #     """ https://docs.viktor.ai/docs/guides/topics/views/create-image-view """
    #     png_path = Path(__file__).parent / 'viktor-logo.png'
    #     return PNGResult.from_path(png_path)


# Create the initial entity of type 'MyEntityType' with the name "demo" as a starting point for the user.
initial_entities = [
    InitialEntity('MyEntityType', name='Demo', params='my_entity.json')  # predefined entity properties from a .json file.
]
