from time import sleep
from typing import List, Tuple

from win32com.client import CastTo

import constants
from caddoc import CADDoc
from point import Point
from utils import vt_point


class Selection:
    def __init__(self, border, items):
        self.border = border
        self.x, self.y, self.z = border.InsertionPoint
        self.items = items

    def move(self, point1: Point, point2: Point):
        for item in self.items:
            item.Move(vt_point(point1), vt_point(point2))


def sort_drawings(selections: List[Selection]) -> List[List]:
    border_width = 841
    border_height = 594
    sorted_by_x = sorted(selections, key=lambda sel: sel.x)
    db = dict()
    for selection in sorted_by_x:
        y = selection.y
        index = round(y / border_height)
        if index not in db:
            db[index] = [selection]
        else:
            db[index].append(selection)
    result = [db[key] for key in sorted(db.keys(), reverse=True)]
    return result


def reset_selection_sets(selection_sets):
    for index in reversed(range(selection_sets.Count)):
        selection_sets.Item(index).Delete()


def align_borders(dwg: CADDoc):
    distance_x = 900
    distance_y = 660
    margin_y = 20
    start_x = 0
    start_y = 0
    border_name = "Border.A1"
    borders = dwg.get_blockrefs(border_name)
    selections = []
    counter = 1
    for border in borders:
        print(f"Preprocessing {counter}/{len(borders)}")
        # Add delay for processing
        sleep(0.05)
        bottom_left, top_right = border.GetBoundingBox()
        point_btm_left = Point(*bottom_left)
        point2 = Point(*top_right)
        point1 = Point(point_btm_left.x, point_btm_left.y - margin_y, point_btm_left.z)
        entities = dwg.select(constants.acSelectionSetCrossing, point1, point2)
        selections.append(Selection(border, entities))
        counter += 1
    # Sorting
    sorted_drawings = sort_drawings(selections)

    new_y = start_y
    for row in sorted_drawings:
        new_x = start_x
        for selection in row:
            point1 = Point(*selection.border.InsertionPoint)
            point2 = Point(new_x, new_y, 0)
            print(f"Moving from {point1} to {point2}")
            selection.move(point1, point2)
            new_x += distance_x
        new_y -= distance_y


def format_pipe_tag(dwg: CADDoc):
    pass


if __name__ == "__main__":
    drawing = CADDoc()
    drawing.replace_text(r'(.*B\d{1}SRF\d{1})(\(.*\))', r'\g<1>')
    drawing.replace_block("TAG_NUMBER", "pipe_tag")
