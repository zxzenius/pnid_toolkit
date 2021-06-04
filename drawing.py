import re
from collections import defaultdict
from typing import List, Iterator
from uuid import uuid4

import constants
from pathlib import Path

from win32com.client import CastTo, Dispatch
from win32com.client.gencache import EnsureDispatch

import dxf
from point import Point
from utils import vt_int_array, vt_variant_array, vt_point


def get_application(app='AutoCAD', version='', early_binding=True, visible=True):
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


def get_attributes(blockref) -> dict:
    """
    Wrapper of Block.GetAttributes
    Usage: attrs = get_attributes(blockRef)
           attrs['TAG'].TextString = 'hello'
    :param blockref: BlockReference
    :return Dict contends AcadAttributeReference objects
    """
    return {attr.TagString: attr for attr in blockref.GetAttributes()}


def get_dynamic_props(blockref) -> dict:
    return {prop.PropertyName: prop for prop in blockref.GetDynamicBlockProperties()}


class Drawing:
    def __init__(self, app_name='autocad', filepath=None, early_binding=True, load_data=True):
        self.early_binding = early_binding
        self.app = get_application(app=app_name, early_binding=early_binding)
        self.doc = None
        self.blockref_db = defaultdict(list)
        self.load(filepath, load_data)

    def init_db(self):
        self.blockref_db = self.gen_blockref_dict()

    def load(self, filepath=None, load_data=True):
        self.doc = get_document(self.app, filepath)
        print(f"Current File: {self.doc.Name}")
        if load_data:
            self.init_db()

    def reload(self):
        self.init_db()

    def reset_selection_sets(self):
        for index in reversed(range(self.doc.SelectionSets)):
            self.doc.SelectionSets.Item(index).Delete()

    def select(self, mode, point1: Point = None, point2: Point = None, filter_type=None, filter_data=None) -> List:
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

    def _select_by_type_and_name(self, type_name: str, entity_name: str) -> List:
        filter_type = vt_int_array([0, 2])
        filter_data = vt_variant_array([type_name, entity_name])
        return self.select(constants.acSelectionSetAll, filter_type=filter_type, filter_data=filter_data)

    def select_entities_by_name(self, dxf_entity: dxf.Entity, entity_name: str) -> List:
        entities = self._select_by_type_and_name(dxf_entity.type_name, entity_name)
        return [CastTo(entity, dxf_entity.interface) for entity in entities]

    def gen_blockref_dict(self, by_select: bool = True) -> dict:
        print("Indexing blockrefs...")
        counter = 0
        db = defaultdict(list)
        if not self.doc:
            return db

        if not by_select:
            blockrefs = self.iter_blockrefs()
        else:
            blockrefs = self.select_blockrefs()
        for blockref in blockrefs:
            db[blockref.EffectiveName].append(blockref)
            counter += 1

        print(f"Indexing complete, {counter} blockrefs.")
        return db

    def get_blockrefs(self, name: str = None) -> List:
        if name:
            return self.get_blockrefs_by_name(name)

        all_blockrefs = []
        for blockrefs in self.blockref_db.values():
            all_blockrefs.extend(blockrefs)
        return all_blockrefs

    def select_blockrefs(self, name: str = None) -> List:
        # not suit for select anonymous blockref by name, which name is "*U..."
        # "name" is not "EffectiveName"
        if name:
            return self.select_entities_by_name(dxf.BlockRef, name)

        return self.select_entities(dxf.BlockRef)

    def iter_entities(self, dxf_entity: dxf.Entity) -> Iterator:
        for item in self.doc.ModelSpace:
            if item.ObjectName == dxf_entity.object_name:
                ent = CastTo(item, dxf_entity.interface)
                yield ent

    def iter_blockrefs(self) -> Iterator:
        return self.iter_entities(dxf.BlockRef)

    def get_blockrefs_by_name(self, name: str):
        return self.blockref_db.get(name)

    def search_blockrefs(self, name_pattern: str) -> List:
        result = []
        prog = re.compile(name_pattern)
        for effective_name in self.blockref_db:
            if prog.match(effective_name):
                result.extend(self.blockref_db[effective_name])

        return result

    def _select_by_type(self, type_name: str) -> List:
        filter_type = vt_int_array([0])
        filter_data = vt_variant_array([type_name])
        entities = self.select(constants.acSelectionSetAll, filter_type=filter_type, filter_data=filter_data)
        return entities

    def select_entities(self, dxf_entity: dxf.Entity) -> List:
        entities = self._select_by_type(dxf_entity.type_name)
        return [CastTo(entity, dxf_entity.interface) for entity in entities]

    def get_block(self, name: str):
        for block in self.doc.Blocks:
            if block.Name == name:
                return block
        return None

    def select_entities_in_area(self, point1: Point, point2: Point, crossing=True) -> List:
        if crossing:
            mode = constants.acSelectionSetCrossing
        else:
            mode = constants.acSelectionSetWindow
        return self.select(mode, point1, point2)

    def select_all_entities(self) -> List:
        return self.select(constants.acSelectionSetAll)

    def iter_all_texts(self) -> Iterator:
        for blockref in self.select_blockrefs():
            # For AutoCAD 2006, using IAcadBlockReference2
            # block_ref = CastTo(item, 'IAcadBlockReference2')
            # 2007 or higher version, using IAcadBlockReference
            for attribute in blockref.GetAttributes():
                if attribute.TextString:
                    yield attribute

        for m_text in self.select_entities(dxf.MText):
            # m_text = CastTo(m_text, 'IAcadMText')
            yield m_text

        for text in self.select_entities(dxf.Text):
            # text = CastTo(text, 'IAcadText')
            yield text

    def replace_text(self, pattern, replacement):
        # scanning all
        print("Start text replacing.")
        regex = re.compile(pattern)
        counter = 0
        for item in self.iter_all_texts():
            if (result := regex.sub(replacement, item.TextString)) != item.TextString:
                item.TextString = result
                counter += 1

        print(f'Replaced {counter} texts.')

    def replace_block(self, old_block_name, new_block_name):
        print("Start block replacing.")
        counter = 0
        blockrefs = self.blockref_db[old_block_name]
        for old_blockref in blockrefs:
            insertion_point = Point(*old_blockref.InsertionPoint)
            new_blockref = self.doc.ModelSpace.InsertBlock(vt_point(insertion_point), new_block_name, 1, 1, 1, 0, None)
            new_blockref.Rotation = old_blockref.Rotation
            new_blockref.Layer = old_blockref.Layer
            old_attributes = get_attributes(old_blockref)
            new_attributes = get_attributes(new_blockref)
            for tag in new_attributes:
                if tag in old_attributes:
                    new_attributes[tag].TextString = old_attributes[tag].TextString # noqa
            old_blockref.Delete()
            counter += 1

        print(f"Replaced {counter} blockrefs.")

    def replace_blockref(self, old_blockref, new_block_name):
        insertion_point = Point(*old_blockref.InsertionPoint)
        new_blockref = self.doc.ModelSpace.InsertBlock(vt_point(insertion_point), new_block_name, 1, 1, 1, 0, None)
        return new_blockref


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
