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


def extract_attributes(blockref) -> dict:
    return {attr.TagString.lower(): attr.TextString for attr in blockref.GetAttributes()}


def extract_dynamic_properties(blockref) -> dict:
    return {prop.PropertyName.lower(): prop.Value for prop in blockref.GetDynamicBlockProperties()}


def is_in_box(point: Point, bottom_left: Point, top_right: Point) -> bool:
    """
    Return true if input point is in the rectangle defined by bottom_left & top_right
    :param point:
    :param bottom_left:
    :param top_right:
    :return: boolean
    """
    return (bottom_left.x < point.x < top_right.x) and (bottom_left.y < point.y < top_right.y)