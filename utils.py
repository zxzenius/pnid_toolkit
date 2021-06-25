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


def get_attributes(blockref) -> dict:
    """
    Wrapper of Block.GetAttributes
    Usage: attrs = get_attributes(blockRef)
           attrs['TAG'].TextString = 'hello'
    :param blockref: BlockReference
    :return Dict contends AcadAttributeReference objects
    """
    return {attr.TagString: attr for attr in blockref.GetAttributes()}


def get_attribute(blockref, tag: str):
    for attr in blockref.GetAttributes():
        if attr.TagString == tag:
            return attr


def get_dynamic_props(blockref) -> dict:
    return {prop.PropertyName: prop for prop in blockref.GetDynamicBlockProperties()}


def get_dynamic_prop(blockref, prop_name: str):
    for prop in blockref.GetDynamicBlockProperties():
        if prop.PropertyName == prop_name:
            return prop
