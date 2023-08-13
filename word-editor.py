from pathlib import Path

import PySimpleGUI as sg
from docxtpl import DocxTemplate

document_path = Path(__file__).parent / "template.docx"
doc = DocxTemplate(document_path)

layout = [
    [sg.Text("Name of project"), sg.Input(key="NAME", do_not_clear=False)],
    [sg.Text("Province(eg. Gandaki)"), sg.Input(key="PROVINCE", do_not_clear=False)],
    [sg.Text("Municipality(eg. Kaligandaki Municipality)"), sg.Input(key="MUNICIPALITY", do_not_clear=False)],
    [sg.Text("Ward no."), sg.Input(key="WARD", do_not_clear=False)],
    [sg.Text("District"), sg.Input(key="DISTRICT", do_not_clear=False)],
    [sg.Text("District Headquarter"), sg.Input(key="HEADQUARTER", do_not_clear=False)],
    [sg.Text("Nearest road"), sg.Input(key="NEARROAD", do_not_clear=False)],
    [sg.Text("Nearest road distance(eg. 5)"), sg.Input(key="NEARROADDISTANCE", do_not_clear=False)],
    [sg.Text("Nearest market"), sg.Input(key="NEARMARKET", do_not_clear=False)],
    [sg.Text("Nearest market distance(eg. 10)"), sg.Input(key="NEARMARKETDISTANCE", do_not_clear=False)],
    [sg.Text("Population"), sg.Input(key="POPULATION", do_not_clear=False)],
    [sg.Text("Household"), sg.Input(key="HOUSEHOLD", do_not_clear=False)],
    [sg.Text("Gross Area(eg. 40)"), sg.Input(key="GROSSAREA", do_not_clear=False)],
    [sg.Text("Net Area(eg. 36)"), sg.Input(key="NETAREA", do_not_clear=False)],
    [sg.Text("Latitude(eg. 27°55'40.31\"N)"), sg.Input(key="LATITUDE", do_not_clear=False)],
    [sg.Text("Longitude(eg.83°41'39.54\"E )"), sg.Input(key="LONGITUDE", do_not_clear=False)],
    [sg.Text("Recommendation"), sg.Multiline(key="RECOMMENDATION", do_not_clear=False)],
    [sg.Text("Enter file name"), sg.Input(key="FILENAME", do_not_clear=False)],

    [sg.Button("Create File"), sg.Exit()],
]

window = sg.Window("Word file generator", layout, element_justification="right")

while True:
    event, values = window.read()
    if event == sg.WIN_CLOSED or event == "Exit":
        break
    if event == "Create File":
        values["FORESTAREA"] = round(float(values["GROSSAREA"]))-round(float(values["NETAREA"]))
        doc.render(values)
        output_path = Path(__file__).parent / f"{values['FILENAME']}.docx"
        doc.save(output_path)
        sg.popup("File created")


# context = {"NAME": "Satunga" }
# doc.render(context)
# doc.save(Path(__file__).parent / "generated.docx")