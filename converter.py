from pypdf import PdfReader
from datetime import datetime
import re
import os
import shutil
import win32com.client
import win32timezone

# ============================================================
# FUNÇÕES AUXILIARES
# ============================================================

def get_file_path(message):
    """
    Pede um caminho até que o utilizador informe
    um ficheiro existente.
    """

    while True:

        path = input(message).strip().strip('"')

        if os.path.isfile(path):
            return path

        print("\nERRO: O ficheiro informado não foi encontrado.")
        print("Tente novamente.\n")


def convert_number(value):

    try:
        number = float(value)

        if number.is_integer():
            return int(number)

        return number

    except (ValueError, TypeError):
        return value


# ============================================================
# EXTRAÇÃO DAS MEDIÇÕES
# ============================================================

def extract_measurements(text):

    measurements = []
    last_measurement = None

    for line in text.splitlines():

        match = re.match(
            r"^\s*(\d+)\s+"
            r"([A-L])\s+"
            r"(\d+(?:\.\d+)?)\s+"
            r"([±+\-]\d+(?:\.\d+)?)\s+"
            r"(.+?)\s*$",
            line
        )

        if match:

            item = int(match.group(1))
            instrument = match.group(2)
            dimension = match.group(3)
            tolerance = match.group(4)

            measures_text = match.group(5)
            measures = measures_text.split()

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
            r"^\s*([+\-]\d+(?:\.\d+)?)\s*$",
            line
        )

        if tolerance_match and last_measurement is not None:

            second_tolerance = tolerance_match.group(1)

            last_measurement["tolerance"].append(
                second_tolerance
            )

    return measurements


# ============================================================
# CONVERTER UM RELATÓRIO
# ============================================================

def convert_report(pdf_path, template_path):

    # --------------------------------------------------------
    # LER PDF
    # --------------------------------------------------------

    report = PdfReader(pdf_path)
    page = report.pages[0]

    text = page.extract_text(
        extraction_mode="layout"
    )


    # --------------------------------------------------------
    # PART NO. ORIGINAL
    # --------------------------------------------------------

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


    # --------------------------------------------------------
    # TAMANHO
    # --------------------------------------------------------

    size_match = re.search(
        r"([0-9]+(?:\s*[\*xX×]\s*[0-9]+))",
        text
    )

    size = (
        size_match.group(1)
        if size_match
        else None
    )


    # --------------------------------------------------------
    # QUANTIDADE
    # --------------------------------------------------------

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


    # --------------------------------------------------------
    # MEDIÇÕES
    # --------------------------------------------------------

    measurements = extract_measurements(text)


    # --------------------------------------------------------
    # MOSTRAR O QUE FOI ENCONTRADO
    # --------------------------------------------------------

    print("\n----------------------------------------")
    print("RELATÓRIO IDENTIFICADO")
    print("----------------------------------------")

    print(f"Part NO. original : {original_part_no}")
    print(f"Tamanho           : {size}")
    print(f"Quantidade        : {quantity}")
    print(f"Medições          : {len(measurements)}")


    # --------------------------------------------------------
    # INPUTS MANUAIS
    # --------------------------------------------------------

    print("\n----------------------------------------")
    print("DADOS DO RELATÓRIO FINAL")
    print("----------------------------------------")

    final_part_no = input(
        "Digite o Part NO. final: "
    ).strip()

    company_name = input(
        "Digite o nome da empresa: "
    ).strip()

    description = input(
        "Digite a descrição: "
    ).strip()


    # --------------------------------------------------------
    # DATA
    # --------------------------------------------------------

    today = datetime.now()


    # --------------------------------------------------------
    # ORGANIZAR DADOS
    # --------------------------------------------------------

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


    # --------------------------------------------------------
    # NOME DO ARQUIVO
    # --------------------------------------------------------

    file_name = (
        f"{original_part_no}-{final_part_no}"
    )


    # ========================================================
    # PASTA DE DESTINO
    # ========================================================

    # Os ficheiros finais serão criados
    # na mesma pasta do relatório original

    output_folder = os.path.dirname(pdf_path)


    # Usa a mesma extensão do template (.xls ou .xlsx)

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


    # ========================================================
    # ABRIR EXCEL
    # ========================================================

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


        # ====================================================
        # DADOS GERAIS
        # ====================================================

        sheet.Range("C3").Value = (
            report_data["company_name"]
        )

        sheet.Range("F3").Value = (
            report_data["description"]
        )

        sheet.Range("C4").Value = (
            report_data["final_part_no"]
        )

        sheet.Range("F4").Value = (
            report_data["size"].replace("*", "x")
        )

        sheet.Range("M4").Value = (
            report_data["quantity"]
        )


        # ====================================================
        # DATA
        # ====================================================

        sheet.Range("C25").Value = today
        sheet.Range("E25").Value = today


        # ====================================================
        # LIMPAR TABELA ANTIGA
        # ====================================================

        sheet.Range(
            "A7:P16"
        ).ClearContents()


        # ====================================================
        # PREENCHER MEDIÇÕES
        # ====================================================

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


        # ====================================================
        # SALVAR EXCEL
        # ====================================================

        workbook.Save()


        # ====================================================
        # GERAR PDF
        # ====================================================

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


# ============================================================
# PROGRAMA PRINCIPAL
# ============================================================

print("========================================")
print("     CONVERSOR DE RELATÓRIOS")
print("========================================")


# ------------------------------------------------------------
# TEMPLATE
# ------------------------------------------------------------

template_path = get_file_path(
    "\nDigite o caminho do template Excel: "
)


print("\nTemplate selecionado:")
print(template_path)


# ============================================================
# LOOP DE CONVERSÕES
# ============================================================

while True:

    print("\n========================================")
    print("NOVO RELATÓRIO")
    print("========================================")


    pdf_path = get_file_path(
        "\nDigite o caminho do relatório original: "
    )


    try:

        convert_report(
            pdf_path,
            template_path
        )

    except Exception as error:

        print("\nERRO DURANTE A CONVERSÃO:")
        print(error)


    # ========================================================
    # OUTRO RELATÓRIO?
    # ========================================================

    print("\n========================================")

    option = input(
        "Deseja converter outro relatório? (S/N): "
    ).strip().lower()


    if option not in ["s", "sim"]:

        break


# ============================================================
# FINAL
# ============================================================

print("\nPrograma encerrado.")