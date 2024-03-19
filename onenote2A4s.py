from PyPDF2 import PdfFileReader, PdfFileWriter
from PyPDF2.pdf import RectangleObject
from PyPDF2.generic import FloatObject
from copy import copy


# ratio from mm to pdf units (points) is approx 2.83465
mm_to_pt_ratio = 2.8346567

def crop_pdf(input_path, output_path, page_number, left=None, right=None, up=None, down=None):
    with open(input_path, 'rb') as file:
        pdf_reader = PdfFileReader(file)
        pdf_writer = PdfFileWriter()

        for page_index in range(pdf_reader.numPages):
            page = pdf_reader.getPage(page_index)
            if page_index == page_number:
                if left is not None or right is not None or up is not None or down is not None:
                    crop_box = page.cropBox.getLowerLeft()
                    upper_right = list(page.cropBox.getUpperRight())
                    
                    if left is not None:
                        crop_box[0] = left * mm_to_pt_ratio
                    if right is not None:
                        upper_right[0] = right * mm_to_pt_ratio
                    if up is not None:
                        upper_right[1] = up * mm_to_pt_ratio
                    if down is not None:
                        crop_box[1] = down * mm_to_pt_ratio
                        
                    page.cropBox.lowerLeft = crop_box
                    page.cropBox.upperRight = upper_right
                    page.mediaBox.lowerLeft = crop_box
                    page.mediaBox.upperRight = upper_right
            pdf_writer.addPage(page)

        with open(output_path, 'wb') as output_file:
            pdf_writer.write(output_file)

def cutA4(input_path, output_path):
    
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


    pdf_in = PdfFileReader(input_path)
    pdf_out = PdfFileWriter()

    print('cutting into pages...')
    for pn in range(pdf_in.getNumPages()):
        page = pdf_in.getPage(pn)
        page_width_pt = page.mediaBox.getUpperRight_x()
        page_height_pt = page.mediaBox.getUpperRight_y()

        scale = float(crop_width_pt / page_width_pt)
        page.scale(scale, scale)
        scale = FloatObject(scale)
        page_width_pt *= scale
        page_height_pt *= scale

        # Calculate the number of parts needed for the current page
        num_of_parts = int(page_height_pt / crop_height_pt) + 1

        for i in range(num_of_parts):
            new_page = copy(page)
            new_page.cropBox = RectangleObject([0, 0, crop_width_pt, crop_height_pt]) #important!
            new_page.cropBox.lowerLeft = (0, page_height_pt - (i + 1) * crop_height_pt - blank_height_low_pt)
            new_page.cropBox.upperRight = (crop_width_pt, page_height_pt - i * crop_height_pt + blank_height_up_pt)
            pdf_out.addPage(new_page)

    output_file_path = output_path.format('output')
    with open(output_file_path, 'wb') as out_file:
        print('writing...')
        pdf_out.write(out_file)
    print('finished')

if __name__ == '__main__':
    # pdf input and output file paths
    pdf_input_path = r'E:\Python\project\pdf editor\homework.pdf'
    pdf_output_path_template = r'e:\Python\project\pdf editor\output_template.pdf'
    pdf_output_path = r'e:\Python\project\pdf editor\output.pdf'

    crop_pdf(pdf_input_path, pdf_output_path_template,0, up = 1244.9)
    cutA4(pdf_output_path_template, pdf_output_path)


