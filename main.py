# extract data form an excel file

import openpyxl
import openpyxl.worksheet

wb = openpyxl.load_workbook(
        "C:/Users/Julian/OneDrive - Quadraam/Biologie begrippen.xlsx"
    )
def open_excel() -> openpyxl.worksheet:
    # Load the workbook
    global wb

    # Get the sheet
    sheet = wb["Blad1"]

    return sheet


def select_sheet(sheet: openpyxl.worksheet) -> tuple:# de biolgiepagina begrippen uit excel naar een lijst
    Begrippen = []
    definities = []
    for cell in sheet["C"]:
        # remove the empty cells
        if cell.value is not None:
            Begrippen.append(cell.value.lower())
    for cell in sheet["D"]:
        if cell.value is not None:

            definities.append(cell.value)
    list(set(Begrippen))

    return Begrippen, definities


def deelconcepten() -> list:

    with open(
        "C:/Users/Julian/OneDrive - Quadraam/deelconcepten.txt", "r", encoding="utf-8"
    ) as file:
        data = file.read()
    data = data.replace("\n", "")

    data_list = [item.strip().lower() for item in data.split(",")]

    list(set(data_list))

    return data_list


def save_concepten_to_Excel(sheet: openpyxl.worksheet, concepten: list, collum: str):
    # Save the concepten to the Excel file
    global wb

    # Get the sheet
    sheet = wb["Blad1"]

    # Write the concepten to the Excel file and begin at row 8
    row = 8
    for concept in concepten:
        sheet[f"{collum}{row}"] = concept
        row += 1

    # Save the workbook
    wb.save(
        "C:/Users/Julian/OneDrive - Quadraam/Biologie begrippen.xlsx"
    )


    
#save_concepten_to_Excel(open_excel(), deelconcepten(), "F")
# save_concepten_to_Excel(open_excel(), select_sheet(open_excel())[0], "H")

def compare(begrips, concepts, definities):
    new_list = []
    def_list = []
    i = 0
    for item in begrips:
        if item.lower() in concepts:
            new_list.append(item.lower())
            def_list.append(definities[i])
        else:
            new_list.append("")
            def_list.append("")
        
            
        
        
        
        i += 1
        

    return new_list, def_list



begrips = select_sheet(open_excel())[0]
concepts = deelconcepten()
definities = select_sheet(open_excel())[1]

print(len(concepts))

save_concepten_to_Excel(open_excel(), compare(begrips, concepts, definities)[0], "F")
save_concepten_to_Excel(open_excel(), compare(begrips, concepts, definities)[1], "G")
