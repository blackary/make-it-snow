from datetime import date
from typing import Any, Dict

from pptx import Presentation
from pptx.shapes.autoshape import Shape
from pptx.slide import Slide

ppt = Presentation("template.pptx")


def describe_slide(slide: Slide):
    for idx, shape in enumerate(slide.shapes):
        print(idx, shape)
        if hasattr(shape, "text"):
            print("\t", shape.text)

        if hasattr(shape, "shapes"):
            for idx2, shape2 in enumerate(shape.shapes):
                print("\t\t", idx2, shape2)
                if hasattr(shape2, "text"):
                    print("\t\t\t", shape2.text)


slide = ppt.slides[0]

shapes = list(slide.shapes)

DATE = shapes[9]

WESTERN_NA = list(shapes[3].shapes)[1]

EUROPE = list(shapes[5].shapes)[1]

AUSTRALIA_EAST = list(shapes[6].shapes)[1]

SINGAPORE = list(shapes[7].shapes)[1]

CENTRAL_NA = list(shapes[8].shapes)[1]

JAPAN = shapes[11]

EASTERN_NA = shapes[13]

INDIA = shapes[15]

ISRAEL = shapes[17]

PHILIPPINES = shapes[19]

NEW_ZEALAND = shapes[21]

AUSTRALIA_WEST = shapes[23]

INDONESIA = list(shapes[26].shapes)[1]

UAE = shapes[25]


def edit_text(shape: Shape, new_text: Any):
    runs = shape.text_frame.paragraphs[0].runs
    for run in runs:
        run.text = ""
    runs[0].text = str(new_text)


def populate_slide(buckets: Dict[str, int], date: date):
    date_str = date.strftime("%B %d, %Y")

    edit_text(DATE, date_str)

    edit_text(AUSTRALIA_WEST, buckets.pop("AUSTRALIA_WEST"))

    edit_text(AUSTRALIA_EAST, buckets.pop("AUSTRALIA_EAST"))

    edit_text(EASTERN_NA, buckets.pop("EASTERN_NA"))

    edit_text(EUROPE, buckets.pop("EUROPE"))

    edit_text(INDIA, buckets.pop("INDIA"))

    edit_text(ISRAEL, buckets.pop("ISRAEL"))

    edit_text(UAE, buckets.pop("UAE"))

    edit_text(WESTERN_NA, buckets.pop("WESTERN_NA"))

    edit_text(CENTRAL_NA, buckets.pop("CENTRAL_NA"))

    edit_text(INDONESIA, buckets.pop("INDONESIA"))

    edit_text(PHILIPPINES, buckets.pop("PHILIPPINES"))

    edit_text(NEW_ZEALAND, buckets.pop("NEW_ZEALAND"))

    edit_text(SINGAPORE, buckets.pop("SINGAPORE_MALAYSIA"))

    edit_text(JAPAN, buckets.pop("JAPAN_SOUTH_KOREA"))

    assert not buckets, buckets

    ppt.save("output.pptx")
