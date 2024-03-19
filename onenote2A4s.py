from PyPDF2 import PdfFileReader, PdfFileWriter
from PyPDF2.pdf import RectangleObject
from PyPDF2.generic import FloatObject
from copy import copy

# pdf input and output file paths
pdf_input_path = r'E:\Python\project\pdf editor\homework.pdf'
pdf_output_path_template = r'e:\Python\project\pdf editor\output_{}.pdf'


def cutA4():
    # ratio from mm to pdf units (points) is approx 2.83465
    mm_to_pt_ratio = 2.83465
    # desired crop size in mm
    crop_width_mm = 210
    crop_height_mm = 290
    blank_height_up_mm =3.5
    blank_height_low_mm =3.5

    # convert from mm to pdf units
    crop_width_pt = FloatObject(crop_width_mm * mm_to_pt_ratio)
    crop_height_pt = FloatObject(crop_height_mm * mm_to_pt_ratio)
    blank_height_up_pt = FloatObject(blank_height_up_mm * mm_to_pt_ratio)
    blank_height_low_pt = FloatObject(blank_height_low_mm * mm_to_pt_ratio)


    pdf_in = PdfFileReader(pdf_input_path)
    pdf_out = PdfFileWriter()

    print('cutting into pages...')
    for pn in range(pdf_in.getNumPages()):
        page = pdf_in.getPage(pn)
        page_width_pt = page.mediaBox.getUpperRight_x()
        page_height_pt = page.mediaBox.getUpperRight_y()

        scale = float(crop_width_pt / page_width_pt)
        page.scale(scale, scale)
        page_width_pt *= FloatObject(scale)
        page_height_pt *= FloatObject(scale)

        # Calculate the number of parts needed for the current page
        num_of_parts = int(page_height_pt / crop_height_pt) + 1

        for i in range(num_of_parts):
            new_page = copy(page)
            new_page.mediaBox = RectangleObject([0, 0, crop_width_pt, crop_height_pt]) #important!
            new_page.mediaBox.lowerLeft = (0, page_height_pt - (i + 1) * crop_height_pt - blank_height_low_pt)
            new_page.mediaBox.upperRight = (crop_width_pt, page_height_pt - i * crop_height_pt + blank_height_up_pt)
            pdf_out.addPage(new_page)

    output_file_path = pdf_output_path_template.format('output')
    with open(output_file_path, 'wb') as out_file:
        print('writing...')
        pdf_out.write(out_file)
    print('finished')

cutA4()