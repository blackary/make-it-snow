from datetime import date
from typing import Any, Dict

from pptx import Presentation
from pptx.shapes.autoshape import Shape

ppt = Presentation("template.pptx")

slide = ppt.slides[0]

shapes = list(slide.shapes)

DATE = shapes[9]

EAST_ASIA = shapes[11]

EASTERN_NA = shapes[13]

WESTERN_NA = list(shapes[3].shapes)[1]

EUROPE = list(shapes[5].shapes)[1]

AUSTRALIA = list(shapes[6].shapes)[1]

SOUTHEAST_ASIA = list(shapes[7].shapes)[1]

CENTRAL_NA = list(shapes[8].shapes)[1]

INDIA = shapes[16]


def edit_text(shape: Shape, new_text: Any):
    shape.text_frame.paragraphs[0].runs[0].text = str(new_text)


def populate_slide(buckets: Dict[str, int], date: date):
    date_str = date.strftime("%B %d, %Y")

    edit_text(DATE, date_str)

    edit_text(AUSTRALIA, buckets["AUSTRALIA"])

    edit_text(EAST_ASIA, buckets["EAST_ASIA"])

    edit_text(EASTERN_NA, buckets["EASTERN_NA"])

    edit_text(EUROPE, buckets["EUROPE"])

    edit_text(INDIA, buckets["INDIA"])

    edit_text(SOUTHEAST_ASIA, buckets["SOUTHEAST_ASIA"])

    edit_text(WESTERN_NA, buckets["WESTERN_NA"])

    edit_text(CENTRAL_NA, buckets["CENTRAL_NA"])

    ppt.save("output.pptx")
