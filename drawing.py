import re
from collections import defaultdict
from typing import List, Generator
from uuid import uuid4

import constants
from pathlib import Path

from win32com.client import CastTo, Dispatch
from win32com.client.gencache import EnsureDispatch

import dxf_names
from point import Point
from utils import vt_int_array, vt_variant_array, vt_point


def get_application(app='AutoCAD', version='', early_binding=False, visible=True):
    """
    AutoCAD Application COM Object
    Also work for BricsCAD
    :param early_binding:
    :param app:
        app name of 'AutoCAD' or 'BricsCAD', case insensitive
        default is 'AutoCAD'
    :param version:
        default is latest version
        '16' for version 2006
    :param visible:
        default is True
    :return:
        AutoCAD Application Object
    """
    if app.lower() == 'autocad':
        prog_id = 'AutoCAD.Application'
    elif app.lower() == 'bricscad':
        prog_id = 'BricscadApp.AcadApplication'
    else:
        raise ValueError('app should be "AutoCAD" or "Bricscad"')
    if version:
        prog_id = '.'.join((prog_id, version))
    if early_binding:
        app = EnsureDispatch(prog_id)
    else:
        app = Dispatch(prog_id)
    app.Visible = visible
    return app


def get_document(app, filename=None):
    # load current file
    if filename is None:
        return app.ActiveDocument

    for document in app.Documents:
        if Path(document.FullName) == Path(filename):
            return document

    return app.Documents.Open(filename)


def get_attributes(block_ref) -> dict:
    """
    Wrapper of Block.GetAttributes
    Usage: attrs = get_attributes(blockRef)
           attrs['TAG'].TextString = 'hello'
    :param block_ref: BlockReference
    :return Dict contends AcadAttributeReference objects
    """
    return {attr.TagString: attr for attr in block_ref.GetAttributes()}


def get_dynamic_props(block_ref) -> dict:
    return {prop.PropertyName: prop for prop in block_ref.GetDynamicBlockProperties()}


class Drawing:
    def __init__(self, app_name='autocad', filepath=None):
        self.app = get_application(app=app_name)
        self.doc = get_document(self.app, filepath)
        self.model_space = self.doc.ModelSpace
        self.block_ref_dict = self.walk_block_refs()

    def clear_selections(self):
        for index in reversed(range(self.doc.SelectionSets)):
            self.doc.SelectionSets.Item(index).Delete()

    def select(self, mode, point1: Point = None, point2: Point = None, filter_type=None, filter_data=None):
        selection_set = self.doc.SelectionSets.Add(uuid4().hex)
        if point1 and point2:
            point1 = vt_point(point1)
            point2 = vt_point(point2)
        if filter_data and filter_data:
            selection_set.Select(mode, point1, point2, filter_type, filter_data)
        else:
            selection_set.Select(mode, point1, point2)
        entities = list(selection_set)
        selection_set.Delete()
        return entities

    def select_by_type_and_name(self, type_name: str, entity_name: str):
        filter_type = vt_int_array([0, 2])
        filter_data = vt_variant_array([type_name, entity_name])
        entities = self.select(constants.acSelectionSetAll, filter_type=filter_type, filter_data=filter_data)
        return entities

    def walk_block_refs(self, brutal_method: bool = False) -> dict:
        print("Indexing block_refs...")
        block_refs = defaultdict(list)
        if brutal_method:
            for entity in self.model_space:
                if entity.ObjectName == "AcDbBlockReference":
                    entity = CastTo(entity, "IAcadBlockReference")
                    real_name = entity.EffectiveName
                    block_refs[real_name].append(entity)
        else:
            for entity in self.get_block_refs():
                entity = CastTo(entity, "IAcadBlockReference")
                real_name = entity.EffectiveName
                block_refs[real_name].append(entity)

        print("Indexing complete.")
        return block_refs

    def regen_block_ref_map(self):
        self.block_ref_dict = self.walk_block_refs()

    def get_all_block_refs(self) -> List:
        return self.select_by_type(dxf_names.INSERT)

    def get_block_refs(self, name: str = None) -> List:
        if name:
            return self.block_ref_dict.get(name)
        return self.select_by_type(dxf_names.INSERT)

    def find_block_refs(self, name_pattern: str) -> List:
        result = []
        prog = re.compile(name_pattern)
        for effective_name in self.block_ref_dict:
            if prog.match(effective_name):
                result.extend(self.block_ref_dict[effective_name])

        return result

    def select_by_type(self, type_name: str):
        filter_type = vt_int_array([0])
        filter_data = vt_variant_array([type_name])
        entities = self.select(constants.acSelectionSetAll, filter_type=filter_type, filter_data=filter_data)
        return entities

    def get_block(self, name: str):
        counter = 0
        for block in self.doc.Blocks:
            counter += 1
            if block.Name == name:
                return block
        # return self.select_by_type_and_name("BLOCK", name)
        return None

    def get_entities(self, point1: Point = None, point2: Point = None, crossing=True):
        if point1 and point2:
            if crossing:
                mode = constants.acSelectionSetCrossing
            else:
                mode = constants.acSelectionSetWindow
            return self.select(mode, point1, point2)
        else:
            mode = constants.acSelectionSetAll
            return self.select(mode)

    def get_all_text(self) -> Generator:
        for block_ref in self.select_by_type(dxf_names.INSERT):
            # For AutoCAD 2006, using IAcadBlockReference2
            # block_ref = CastTo(item, 'IAcadBlockReference2')
            # 2007 or higher version, using IAcadBlockReference
            block_ref = CastTo(block_ref, 'IAcadBlockReference')
            for attribute in block_ref.GetAttributes():
                if attribute.TextString:
                    yield attribute

        for m_text in self.select_by_type(dxf_names.MTEXT):
            m_text = CastTo(m_text, 'IAcadMText')
            yield m_text

        for text in self.select_by_type(dxf_names.TEXT):
            text = CastTo(text, 'IAcadText')
            yield text

    def replace_text(self, pattern, replacement):
        # scanning all
        print("Start text replacing.")
        regex = re.compile(pattern)
        counter = 0
        for item in self.get_all_text():
            if (result := regex.sub(replacement, item.TextString)) != item.TextString:
                item.TextString = result
                counter += 1

        print(f'Replaced {counter} texts.')

    def replace_block(self, old_block_name, new_block_name):
        print("Start block replacing.")
        counter = 0
        block_refs = self.get_block_refs(old_block_name)
        for old_block_ref in block_refs:
            insertion_point = Point(*old_block_ref.InsertionPoint)
            new_block_ref = self.model_space.InsertBlock(vt_point(insertion_point), new_block_name, 1, 1, 1, 0, None)
            new_block_ref.Rotation = old_block_ref.Rotation
            new_block_ref.Layer = old_block_ref.Layer
            old_attributes = get_attributes(old_block_ref)
            new_attributes = get_attributes(new_block_ref)
            for tag in new_attributes:
                if tag in old_attributes:
                    new_attributes[tag].TextString = old_attributes[tag].TextString
            old_block_ref.Delete()
            counter += 1

        print(f"Replaced {counter} block_refs.")

    def replace_block_ref(self, old_block_ref, new_block_name):
        insertion_point = Point(*old_block_ref.InsertionPoint)
        new_block_ref = self.model_space.InsertBlock(vt_point(insertion_point), new_block_name, 1, 1, 1, 0, None)
        return new_block_ref


if __name__ == "__main__":
    dwg = Drawing()
    # borders = dwg.get_block_refs("Border.A1")
    # border_pos = [border.InsertionPoint for border in borders]
    # print(border_pos)
    # print(dwg.get_block("TAG_NUMBER"))
    # blocks = dwg.walk_blocks(True)
    # print(len(dwg.block_ref_dict.get("VAL_RELIEF")))
    # print(len(dwg.find_block_refs(r"VAL_.*")))
    # print(len(dwg.get_block_refs()))
    # print(len(dwg.select_by_type("INSERT")))
    # print(len(dwg.get_block_refs("Border*")))
    dwg.replace_block("TAG_NUMBER", "pipe_tag")