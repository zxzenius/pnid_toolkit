from typing import List

from win32com.client import VARIANT
import pythoncom as p

from point import Point


def vt_int_array(values: List[int]) -> VARIANT:
    return VARIANT(p.VT_ARRAY | p.VT_I2, values)


def vt_variant_array(values: List) -> VARIANT:
    return VARIANT(p.VT_ARRAY | p.VT_VARIANT, values)


def vt_point(point: Point) -> VARIANT:
    return VARIANT(p.VT_ARRAY | p.VT_R8, (point.x, point.y, point.z))
