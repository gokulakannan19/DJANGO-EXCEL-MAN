from os import write
from django.shortcuts import render, redirect
from django.http import HttpResponse
from .forms import DocumentForm
from .models import Document
import pandas as pd
import openpyxl
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Border, Side, Alignment, Protection, Font


def home(request):

    documents = Document.objects.all()

    form = DocumentForm()

    if request.method == "POST":
        form = DocumentForm(request.POST, request.FILES)
        if form.is_valid():
            form.save()
            return redirect('home')

    context = {
        'documents': documents,
        'form': form
    }

    return render(request, 'impactgenerator/home.html', context)


def create_impact(request, pk):
    document = Document.objects.get(id=pk)
    file = document.document
    name = document.name

    read_wb = openpyxl.load_workbook(file)
    print(read_wb)

    response = HttpResponse(
        content_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
    response['Content-Dispositon'] = 'attachment; filename="ImpactAnalysis.xlsx"'

    write_wb = Workbook()
    write_wb['Sheet'].title = "Impact Analysis"
    write_sheet = write_wb.active

    write_sheet.column_dimensions['B'].width = 20
    write_sheet.column_dimensions['C'].width = 20
    write_sheet.column_dimensions['D'].width = 50
    write_sheet.column_dimensions['E'].width = 20
    write_sheet.column_dimensions['F'].width = 50
    write_sheet.column_dimensions['G'].width = 50
    write_sheet.column_dimensions['H'].width = 50
    write_sheet.column_dimensions['I'].width = 50
    write_sheet.column_dimensions['J'].width = 50
    write_sheet.column_dimensions['K'].width = 20
    write_sheet.column_dimensions['L'].width = 20
    write_sheet.column_dimensions['M'].width = 20
    write_sheet.column_dimensions['N'].width = 20
    write_sheet.column_dimensions['O'].width = 50

    font1 = Font(name='Calibri', bold=True, size=14)
    font2 = Font(name='Calibri', bold=True, size=11)
    # Heading row 1
    write_sheet['B2'] = "Test Execution on"
    # write_sheet['B2'].font = font1
    write_sheet['C2'] = ""
    # write_sheet['C2'].font = font1
    write_sheet['D2'] = "Executed By"
    # write_sheet['D2'].font = font1
    write_sheet['E2'] = ""
    # write_sheet['E2'].font = font1

    # Heading Row 2
    write_sheet['B3'] = "Test Execution for"
    # write_sheet['B3'].font = font1
    write_sheet['C3'] = ""
    # write_sheet['C3'].font = font1
    write_sheet['D3'] = "Reviewed By"
    # write_sheet['D3'].font = font1
    write_sheet['E3'] = ""
    # write_sheet['E3'].font = font1
    for row in range(2, 4):
        for col in range(1, write_sheet.max_column+1):
            write_sheet.cell(row=row, column=col).font = font1

    write_sheet['A5'] = "S.No"
    write_sheet['B5'] = "Status"
    write_sheet['C5'] = "Bug ID"
    write_sheet['D5'] = "Bug Description"
    write_sheet['E5'] = "DDC"
    write_sheet['F5'] = "Area of Fix from GE(file name)"
    write_sheet['G5'] = "Before Fix"
    write_sheet['H5'] = "Afont1er Fix"
    write_sheet['I5'] = "Scenario description in detail"
    write_sheet['J5'] = "Expected output"
    write_sheet['K5'] = "CH LL review status (For J ~ N col)"
    write_sheet['L5'] = "Retest Result"
    write_sheet['M5'] = "Executed on (DD/MM/YYYY)"
    write_sheet['N5'] = "CH Bug ID"
    write_sheet['O5'] = "Comments"

    for col in range(1, write_sheet.max_column+1):
        write_sheet.cell(row=5, column=col).font = font2

    # to get the current sheet name
    sheet_name = read_wb.active.title
    # print(sheet_name)

    # reading the contents of the current sheet
    sheet1 = read_wb[sheet_name]
    # print(sheet1)

    # to know the maximum countber of columns
    column_end = sheet1.max_column

    row = 1

    for col in range(1, column_end+1):
        # print(row, col)
        column_header = sheet1.cell(row, col).value
        print(column_header)
        # print(len(column_header))
        if column_header == "Status":
            print("found")
            col_value = col
            # print(col_value)
            s_no = 1
            count = 6
            for row_count in range(2, sheet1.max_row+1):
                result = sheet1.cell(row_count, col_value).value

                write_sheet['B'+str(count)] = result
                write_sheet['A'+str(count)] = s_no
                count = count + 1
                s_no = s_no+1
                # print(count)
        elif column_header == "#":
            print("found")
            col_value = col
            # print(col_value)
            count = 6
            for row_count in range(2, sheet1.max_row+1):
                result = sheet1.cell(row_count, col_value).value

                write_sheet['C'+str(count)] = result
                count = count + 1
                # print(count)
        elif column_header == "Description":
            print("found")
            col_value = col
            # print(col_value)
            count = 6
            for row_count in range(2, sheet1.max_row+1):
                result = sheet1.cell(row_count, col_value).value
                # print(result)
                result = result.lower()

                desc_start = result.index("steps")
                desc_string = result[desc_start:]
                desc_end = desc_string.index("\n\n")
                description = desc_string[:desc_end]
                print(description)
                write_sheet['I'+str(count)] = description.strip()
                write_sheet['I'+str(count)
                            ].alignment = Alignment(wrapText=True)
                # print(result.index("expected"))

                expected_result_start = result.index("expected")
                expected_string = result[expected_result_start:]
                expected_result_end = expected_string.index('\n\n')
                expected_result = expected_string[:expected_result_end]
                print(expected_result)
                write_sheet['H'+str(count)] = expected_result.strip()
                write_sheet['H'+str(count)
                            ].alignment = Alignment(wrapText=True)

                write_sheet['J'+str(count)] = expected_result.strip()
                write_sheet['J'+str(count)
                            ].alignment = Alignment(wrapText=True)

                actual_result_start = result.index("actual")
                actual_string = result[actual_result_start:]
                actual_result_end = actual_string.index('\n\n')
                actual_result = actual_string[:actual_result_end]
                print(actual_result)
                write_sheet['G'+str(count)] = actual_result.strip()
                write_sheet['G'+str(count)
                            ].alignment = Alignment(wrapText=True)

                count = count + 1
                # print(count)

        elif column_header == "Details Defect Category":
            print("found")
            col_value = col
            # print(col_value)
            count = 6
            for row_count in range(2, sheet1.max_row+1):
                result = sheet1.cell(row_count, col_value).value

                write_sheet['E'+str(count)] = result.strip()
                count = count + 1

        elif column_header == "Subject":
            print("found")
            col_value = col
            # print(col_value)
            count = 6
            for row_count in range(2, sheet1.max_row+1):
                result = sheet1.cell(row_count, col_value).value

                write_sheet['D'+str(count)] = result.strip()
                count = count + 1

    alignment = Alignment(
        vertical='center',
        wrap_text=True,
    )

    for row in range(1, write_sheet.max_row+1):
        for col in range(1, write_sheet.max_column+1):
            write_sheet.cell(row=row, column=col).alignment = alignment
    write_wb.save(response)
    return response


def delete_document(request, pk):
    document = Document.objects.get(id=pk)

    if request.method == "POST":
        document.delete()
        return redirect('home')

    context = {
        'document': document
    }

    return render(request, 'impactgenerator/delete.html', context)
