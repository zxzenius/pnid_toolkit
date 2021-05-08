from time import sleep
from typing import List, Tuple

from win32com.client import CastTo

import pnid
from utils import vt_int_array, vt_variant_array, vt_point
from uuid import uuid4


acSelectionSetWindow = 0
acSelectionSetCrossing = 1
acSelectionSetAll = 5


class Selection:
    def __init__(self, border, items):
        self.border = border
        self.x, self.y, self.z = border.InsertionPoint
        self.items = items

    def move(self, point1, point2):
        for item in self.items:
            item.Move(vt_point(*point1), vt_point(*point2))


def sort_borders(borders: List) -> List[List]:
    borders_by_x = sorted(borders, key=lambda border: border.InsertionPoint[0])
    db = dict()
    for border in borders_by_x:
        y = border.InsertionPoint[1]
        index = round(y / border_height)
        if index not in db:
            db[index] = [border]
        else:
            db[index].append(border)
    result = [db[key] for key in sorted(db.keys(), reverse=True)]
    return result


def sort_drawings(selections: List[Selection]) -> List[List]:
    sorted_by_x = sorted(selections, key=lambda selection: selection.x)
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


if __name__ == "__main__":
    app = pnid.get_application('AutoCAD')
    doc = pnid.get_document(app)
    border_width = 841
    border_height = 594
    distance_x = 900
    distance_y = 660
    margin_y = 0
    start_x = 0
    start_y = 0
    border_name = "Border.A1"
    block_ref_name = "INSERT"
    filter_type = vt_int_array([2])
    filter_data = vt_variant_array([border_name])
    selection_sets = doc.SelectionSets
    reset_selection_sets(selection_sets)
    s_set = selection_sets.Add(uuid4().hex)
    s_set.Select(acSelectionSetAll, None, None, filter_type, filter_data)
    borders = []
    selections = []
    counter = 1
    for entity in s_set:
        print(f"Preprocessing {counter}/{s_set.Count}")
        border = CastTo(entity, 'IAcadBlockReference')
        sleep(0.02)
        bottom_left, top_right = border.GetBoundingBox()
        x, y, z = bottom_left
        point1 = vt_point(x, y - margin_y, z)
        point2 = vt_point(*top_right)
        new_set = selection_sets.Add(str(counter))
        new_set.Select(acSelectionSetCrossing, point1, point2)
        selections.append(Selection(border, new_set))
        counter += 1
        # borders.append(border)
    # Sorting
    # sorted_borders = sort_borders(selections)
    sorted_drawings = sort_drawings(selections)
    # print(sorted_drawings)
    # print(border_db.keys())
    # sorted_borders = sorted(borders, key=lambda border: round(border.InsertionPoint[1] / border_height))
    # new_y = start_y
    # for row in sorted_borders:
    #     new_x = start_x
    #     for border in row:
    #         pass

    # for s in selection_sets:
    #     s.Delete()

    new_y = start_y
    for row in sorted_drawings:
        new_x = start_x
        for selection in row:
            point1 = selection.border.InsertionPoint
            point2 = (new_x, new_y, 0)
            print(f"Moving from {point1} to {point2}")
            selection.move(point1, point2)
            new_x += distance_x
        new_y -= distance_y

    reset_selection_sets(selection_sets)
