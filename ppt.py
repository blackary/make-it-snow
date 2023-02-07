from datetime import date
from typing import Any, Dict

from pptx import Presentation
from pptx.oxml.shapes.groupshape import CT_GroupShape
from pptx.shapes.autoshape import Shape
from pptx.slide import Slide


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


def edit_text(shape: Shape, new_text: Any):
    # If the text is 0, remove the parent shape
    if str(new_text) == "0":
        parent = shape._element.getparent()
        grandparent = parent.getparent()

        if type(grandparent) == CT_GroupShape:
            grandparent.remove(parent)
        else:
            parent.remove(shape._element)
        return

    runs = shape.text_frame.paragraphs[0].runs

    for run in runs:
        run.text = ""
    to_add = str(new_text)
    if len(to_add) == 1:
        to_add = " " + to_add
    runs[0].text = str(to_add)


def populate_slide(buckets: Dict[str, int], date: date):
    ppt = Presentation("template.pptx")

    slide = ppt.slides[0]

    shapes = list(slide.shapes)

    DATE = shapes[9]

    WESTERN_NA = list(shapes[3].shapes)[1]

    EUROPE = list(shapes[5].shapes)[1]

    AUSTRALIA_EAST = list(shapes[6].shapes)[1]

    SINGAPORE = list(shapes[7].shapes)[1]

    CENTRAL_NA = list(shapes[8].shapes)[1]

    EASTERN_NA = list(shapes[10].shapes)[1]

    ISRAEL = list(shapes[11].shapes)[1]

    INDIA = list(shapes[12].shapes)[1]

    JAPAN = list(shapes[13].shapes)[1]

    PHILIPPINES = list(shapes[14].shapes)[1]

    INDONESIA = list(shapes[15].shapes)[1]

    AUSTRALIA_WEST = list(shapes[16].shapes)[1]

    NEW_ZEALAND = list(shapes[17].shapes)[1]

    UAE = list(shapes[18].shapes)[1]

    BRAZIL = list(shapes[19].shapes)[1]

    COSTA_RICA = list(shapes[20].shapes)[1]

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

    edit_text(BRAZIL, buckets.pop("BRAZIL"))

    edit_text(COSTA_RICA, buckets.pop("COSTA_RICA"))

    assert not buckets, buckets

    ppt.save("output.pptx")
