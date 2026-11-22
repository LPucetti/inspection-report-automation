from pypdf import PdfReader
from datetime import datetime
import re
import os
import shutil
import win32com.client
import win32timezone

def get_file_path(message):

    while True:

        path = input(message).strip().strip('"')

        if os.path.isfile(path):
            return path

        print("\nERRO: O ficheiro informado não foi encontrado.")
        print("Tente novamente.\n")

def get_client_reports():

    pdf_paths = []

    print("\nAdicione os relatórios deste cliente.")
    print("Quando terminar, pressione ENTER sem digitar nenhum caminho.\n")

    while True:

        path = input(
            f"Relatório {len(pdf_paths) + 1}: "
        ).strip().strip('"')

        # ENTER vazio = terminou a seleção
        if path == "":

            if len(pdf_paths) == 0:
                print("\nAdicione pelo menos um relatório.\n")
                continue

            break

        # Verifica se existe
        if not os.path.isfile(path):
            print("ERRO: ficheiro não encontrado.\n")
            continue

        # Verifica se é PDF
        if os.path.splitext(path)[1].lower() != ".pdf":
            print("ERRO: o ficheiro deve ser um PDF.\n")
            continue

        # Evita adicionar o mesmo relatório duas vezes
        if path in pdf_paths:
            print("Este relatório já foi adicionado.\n")
            continue

        pdf_paths.append(path)

        print(
            f"Adicionado: {os.path.basename(path)}"
        )

    return pdf_paths

def convert_number(value):

    try:

        if isinstance(value, str):
            value = value.strip().replace(",", ".")

        number = float(value)

        if number.is_integer():
            return int(number)

        return number

    except (ValueError, TypeError):
        return value
        
def extract_measurements(text):

    measurements = []
    last_measurement = None

    for line in text.splitlines():

        match = re.match(
            r"^\s*(\d+)\s+"
            r"([A-L])\s+"
            r"(\d+(?:[.,]\d+)?)\s+"
            r"([±+\-]\d+(?:[.,]\d+)?)\s+"
            r"(.+?)\s*$",
            line
        )

        if match:

            item = int(match.group(1))
            instrument = match.group(2)

            dimension = match.group(3).replace(",", ".")
            tolerance = match.group(4).replace(",", ".")

            measures_text = match.group(5)

            measures = [
                value.replace(",", ".")
                for value in measures_text.split()
            ]

            measurement = {
                "item": item,
                "instrument": instrument,
                "dimension": dimension,
                "tolerance": [tolerance],
                "measures": measures
            }

            measurements.append(measurement)

            last_measurement = measurement

            continue


        # Segunda linha de tolerância
        tolerance_match = re.match(
            r"^\s*([+\-]\d+(?:[.,]\d+)?)\s*$",
            line
        )

        if tolerance_match and last_measurement is not None:

            second_tolerance = (
                tolerance_match.group(1)
                .replace(",", ".")
            )
            last_measurement["tolerance"].append(
                second_tolerance
            )

    return measurements

def extract_size(text):

    lines = text.splitlines()

    size_header_index = None
    size_start = None
    size_end = None

    for index, line in enumerate(lines):

        match = re.search(
            r"\bSize\b",
            line,
            re.IGNORECASE
        )

        if match:

            size_header_index = index
            size_start = match.start()

            possible_limits = []

            for header_line in lines[
                index:index + 3
            ]:

                for label in [
                    "Inspection",
                    "Quantity",
                    "Re."
                ]:

                    position = header_line.find(
                        label
                    )

                    if position > size_start:
                        possible_limits.append(
                            position
                        )

            if possible_limits:
                size_end = min(
                    possible_limits
                )

            break

    if size_header_index is None:
        return None

    size_parts = []

    for line in lines[
        size_header_index + 1:
    ]:

        if re.search(
            r"\bItem\b",
            line,
            re.IGNORECASE
        ):
            break


        if size_end is not None:

            content = line[
                size_start:size_end
            ].strip()

        else:

            content = line[
                size_start:
            ].strip()


        if not content:
            continue

        if content.lower() in [
            "inspection",
            "quantity",
            "re."
        ]:
            continue


        size_parts.append(content)


    if not size_parts:
        return None

    return "\n".join(size_parts)

def convert_report(pdf_path, template_path,company_name):

    print("\nFicheiro:")
    print(os.path.basename(pdf_path))

    report = PdfReader(pdf_path)
    page = report.pages[0]

    text = page.extract_text(
        extraction_mode="layout"
    )

    part_match = re.search(
        r"Part\s*NO\.\s+.*?\n\s*([A-Za-z0-9._\-]+)",
        text,
        re.IGNORECASE
    )

    original_part_no = (
        part_match.group(1)
        if part_match
        else None
    )

    size=extract_size(text)

    quantity_match = re.search(
        r"Quantity\s+(\d+)",
        text,
        re.IGNORECASE
    )

    quantity = (
        int(quantity_match.group(1))
        if quantity_match
        else None
    )

    measurements = extract_measurements(text)

    print("\n----------------------------------------")
    print("RELATÓRIO IDENTIFICADO")
    print("----------------------------------------")

    print(f"Part NO. original : {original_part_no}")
    print(f"Tamanho           : {size}")
    print(f"Quantidade        : {quantity}")
    print(f"Medições          : {len(measurements)}")

    print("\n----------------------------------------")
    print("DADOS DO RELATÓRIO FINAL")
    print("----------------------------------------")

    final_part_no = input(
        "Digite o Part NO. final: "
    ).strip()

    description = input(
        "Digite a descrição: "
    ).strip()

    today = datetime.now()

    report_data = {

        "original_part_no": original_part_no,

        "final_part_no": final_part_no,

        "company_name": company_name,

        "description": description,

        "size": size,

        "quantity": quantity,

        "date": today,

        "measurements": measurements
    }

    file_name = (
        f"{original_part_no}-{final_part_no}"
    )

    output_folder = os.path.dirname(pdf_path)

    template_extension = os.path.splitext(
        template_path
    )[1]


    output_excel = os.path.join(
        output_folder,
        f"{file_name}{template_extension}"
    )

    output_pdf = os.path.join(
        output_folder,
        f"{file_name}.pdf"
    )


    # ========================================================
    # CRIAR CÓPIA DO TEMPLATE
    # ========================================================

    shutil.copy2(
        template_path,
        output_excel
    )

    excel = None
    workbook = None

    try:

        excel = win32com.client.Dispatch(
            "Excel.Application"
        )

        excel.Visible = False
        excel.DisplayAlerts = False

        workbook = excel.Workbooks.Open(
            output_excel
        )

        sheet = workbook.Worksheets(1)

        sheet.Range("C3").Value = (
            report_data["company_name"]
        )

        sheet.Range("F3").Value = (
            report_data["description"]
        )

        sheet.Range("C4").Value = (
            report_data["final_part_no"]
        )

        size_for_excel = (
            report_data["size"] or ""
        ).replace("*", "x")

        sheet.Range("F4").Value = size_for_excel

        sheet.Range("M4").Value = (
            report_data["quantity"]
        )

        sheet.Range("C25").Value = today
        sheet.Range("E25").Value = today

        sheet.Range(
            "A7:P16"
        ).ClearContents()

        for measurement in report_data["measurements"]:

            item = measurement["item"]

            row = item + 6


            # Item
            sheet.Cells(
                row,
                1
            ).Value = item


            # Instrumento
            sheet.Cells(
                row,
                2
            ).Value = measurement["instrument"]


            # Dimensão
            sheet.Cells(
                row,
                3
            ).Value = convert_number(
                measurement["dimension"]
            )


            # Tolerância

            tolerance_text = "\n".join(
                measurement["tolerance"]
            )

            sheet.Cells(
                row,
                4
            ).Value = tolerance_text


            # Medições

            for index, measure in enumerate(
                measurement["measures"]
            ):

                column = 5 + index

                sheet.Cells(
                    row,
                    column
                ).Value = convert_number(
                    measure
                )

        workbook.Save()

        if os.path.exists(output_pdf):
            os.remove(output_pdf)


        sheet.ExportAsFixedFormat(
            Type=0,
            Filename=output_pdf
        )


        print("\n----------------------------------------")
        print("CONVERSÃO CONCLUÍDA")
        print("----------------------------------------")

        print(
            f"Excel: {output_excel}"
        )

        print(
            f"PDF:   {output_pdf}"
        )


    finally:

        if workbook is not None:
            workbook.Close(
                SaveChanges=False
            )

        if excel is not None:
            excel.Quit()

print("========================================")
print("     CONVERSOR DE RELATÓRIOS")
print("========================================")

template_path = get_file_path(
    "\nDigite o caminho do template Excel: "
)

print("\nTemplate selecionado:")
print(template_path)

while True:

    print("\n========================================")
    print("NOVO CLIENTE")
    print("========================================")

    pdf_paths = get_client_reports()

    print("\n----------------------------------------")
    print("DADOS DO CLIENTE")
    print("----------------------------------------")

    company_name = input(
        "Digite o nome da empresa: "
    ).strip()

    total_reports = len(pdf_paths)

    print("\n----------------------------------------")
    print("CLIENTE CARREGADO")
    print("----------------------------------------")

    print(f"Empresa     : {company_name}")
    print(f"Relatórios  : {total_reports}")


    for index, pdf_path in enumerate(
        pdf_paths,
        start=1
    ):

        print("\n========================================")

        print(
            f"RELATÓRIO {index}/{total_reports}"
        )

        print("========================================")

        print(
            os.path.basename(pdf_path)
        )


        try:

            convert_report(
                pdf_path,
                template_path,
                company_name
            )

        except Exception as error:

            print("\nERRO DURANTE A CONVERSÃO:")

            print(
                os.path.basename(pdf_path)
            )

            print(error)


    print("\n========================================")
    print("CLIENTE CONCLUÍDO")
    print("========================================")

    print(f"Empresa: {company_name}")

    print(
        f"Relatórios processados: {total_reports}"
    )

    option = input(
        "\nDeseja processar outro cliente? (S/N): "
    ).strip().lower()

    if option not in ["s", "sim"]:

        break

print("\nPrograma encerrado.")