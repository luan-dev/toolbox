import os, io, time
import argparse

from pptx import Presentation
from pptx.util import Inches
from pdf2image import convert_from_path

# decorator for timing
def time_it(func):
    def wrapper(*args, **kwargs):
        start_time: float = time.time()
        result = func(*args, **kwargs)
        print(f"{func.__name__} complete! {time.time() - start_time} seconds\n")
        return result

    return wrapper

@time_it
def convert(input_file: str, output_file: str, legacy_res: bool):
    # load file
    pages: list = load_pdf(input_file)

    # create presentation
    ppt = Presentation()
    if not legacy_res:
        ppt.slide_width = Inches(16)
        ppt.slide_height = Inches(9)
    create_ppt(ppt, pages)

    # save presentation
    if not output_file:
        # replace pdf with pptx
        output_file = f"{input_file.split('.', 1)[0]}.pptx"
    ppt.save(output_file)

    # print file stats
    print(f"{len(pages)} pages converted to {output_file}")
    size_in_mb: int = round(os.stat(output_file).st_size / (1024 * 1024), 2)
    print(f"File size: {size_in_mb} MB\n")


# Load PDF
@time_it
def load_pdf(file: str):
    print(f"Loading {file}...")

    # this seems to be the optimal setting without losing too much quality
    pages = convert_from_path(file, dpi=150, fmt="jpeg", thread_count=4)
    return pages


# Crnate slides
@time_it
def create_ppt(ppt: Presentation, pages: list):
    print("Creating Powerpoint...")

    # slide config
    blank_slide_layout = ppt.slide_layouts[6]
    left = top = Inches(0)
    width = ppt.slide_width
    height = ppt.slide_height

    for page in pages:
        # add slide to presentation without saving image
        with io.BytesIO() as output:
            page.save(output, format="JPEG")
            slide = ppt.slides.add_slide(blank_slide_layout)
            slide.shapes.add_picture(output, left, top, width, height)


if __name__ == "__main__":
    # parse arguments
    parser = argparse.ArgumentParser()
    parser.add_argument("input_file", help="PDF file to convert")
    parser.add_argument("-o", "--output", help="output file name (must include extension)", type=str)
    parser.add_argument("-l", "--legacy_res", help="legacy resolution to support 4:3", type=bool)
    args = parser.parse_args()

    # pass arguments to main
    convert(args.input_file, args.output, args.legacy_res)