# Tagging tie-in points
from collections import defaultdict

from pnid import PnID
from point import Point
from utils import get_attribute


def tagging(pnid: PnID):
    tp_counters = defaultdict(int)
    tpoints_by_drawing = defaultdict(list)
    counter = 0
    # Group by drawing
    for tpoint in pnid.blockrefs["TieIn"]:
        drawing = pnid.locate(tpoint)
        tpoints_by_drawing[int(drawing.id[-4:])].append(tpoint)

    for number in sorted(tpoints_by_drawing):

        if number > 1:
            tpoints = tpoints_by_drawing[number]
            sorted_by_y = sorted(tpoints, key=lambda tp: -Point(*tp.InsertionPoint).y)
            sorted_by_x = sorted(sorted_by_y, key=lambda tp: Point(*tp.InsertionPoint).x)
            for tp in sorted_by_x:
                tag = get_attribute(tp, "TAG")
                unit = tag.TextString[:2]
                tp_counters[unit] += 1
                tag.TextString = f"{unit}{tp_counters[unit]:02}"
                counter += 1

    print(f"{counter} tps tagged.")
    print("tpoints_counter:")
    print(tp_counters)


if __name__ == "__main__":
    p = PnID()
    tagging(p)
