import logging
from typing import List

from win32com.client import VARIANT
import pythoncom as p
from win32com.client.gencache import EnsureDispatch

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


def get_dynamic_properties(blockref) -> dict:
    return {prop.PropertyName: prop for prop in blockref.GetDynamicBlockProperties()}


def get_dynamic_property(blockref, prop_name: str):
    for prop in blockref.GetDynamicBlockProperties():
        if prop.PropertyName == prop_name:
            return prop


def copy_attributes(source_bref, target_bref):
    """
    Copy attrs from source_block_ref to target_blockref
    :param source_bref:
    :param target_bref:
    :return:
    """
    old_attrs = get_attributes(source_bref)
    new_attrs = get_attributes(target_bref)
    for tag in new_attrs:
        if tag in old_attrs:
            new_attrs[tag].TextString = old_attrs[tag].TextString
            new_attrs[tag].Alignment = old_attrs[tag].Alignment
            new_attrs[tag].Height = old_attrs[tag].Height
            new_attrs[tag].Layer = old_attrs[tag].Layer
            new_attrs[tag].Rotation = old_attrs[tag].Rotation
            new_attrs[tag].ScaleFactor = old_attrs[tag].ScaleFactor
            new_attrs[tag].StyleName = old_attrs[tag].StyleName
            new_attrs[tag].UpsideDown = old_attrs[tag].UpsideDown
            new_attrs[tag].Visible = old_attrs[tag].Visible
            new_attrs[tag].InsertionPoint = vt_point(Point(*old_attrs[tag].InsertionPoint))


def copy_dynamic_properties(source_bref, target_bref):
    """
    Copy dynamic props from source blockref to target blockref
    :param source_bref:
    :param target_bref:
    :return:
    """
    old_props = get_dynamic_properties(source_bref)
    new_props = get_dynamic_properties(target_bref)
    for name in new_props:
        if name in old_props:
            new_props[name].Value = old_props[name].Value


def get_application(prog_id: str):
    # ref:https://gist.github.com/rdapaz/63590adb94a46039ca4a10994dff9dbe#gistcomment-2918299
    # logger = logging.getLogger(__name__)
    try:
        return EnsureDispatch(prog_id)
    except AttributeError:
        import re
        import sys
        import shutil
        import win32com
        # Remove cache and try again.
        print('Regenerate cache...')
        gen_path = win32com.__gen_path__
        modules = [m.__name__ for m in sys.modules.values()]
        for module in modules:
            if re.match(r'win32com\.gen_py\..+', module):
                del sys.modules[module]
        # Remove gen_py folder
        shutil.rmtree(gen_path)
        # reload
        return win32com.client.gencache.EnsureDispatch(prog_id)
