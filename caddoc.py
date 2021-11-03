import logging
import re
from collections import defaultdict
from pathlib import Path
from typing import List, Iterator, Iterable
from uuid import uuid4

from win32com.client import CastTo

import constants
import dxf
from point import Point
from utils import vt_int_array, vt_variant_array, vt_point, copy_attributes, copy_dynamic_properties, get_application


def get_acad_app(version=''):
    """
    AutoCAD Application COM Object
    :param version:
        default is latest version
        '16' for version 2006
    :return:
        AutoCAD Application Object
    """
    prog_id = 'AutoCAD.Application'
    if version:
        prog_id = '.'.join((prog_id, version))

    return get_application(prog_id)


def get_document(app, filename=None):
    # load current file
    if filename is None:
        return app.ActiveDocument

    for document in app.Documents:
        if Path(document.FullName) == Path(filename):
            return document

    return app.Documents.Open(filename)


class CADDoc:
    def __init__(self, filepath=None):
        self.app = get_acad_app()
        self.doc = None
        self.blockrefs = defaultdict(list)
        # self.logger = logging.getLogger(__name__)
        self.load(filepath)

    def init_db(self):
        self.blockrefs = self.gen_blockref_dict()

    def load(self, filepath=None):
        self.doc = get_document(self.app, filepath)
        print(f"Current File: {self.doc.Name}")
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
        for blockrefs in self.blockrefs.values():
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
        return self.blockrefs.get(name)

    def search_blockrefs(self, name_pattern: str) -> List:
        result = []
        prog = re.compile(name_pattern)
        for effective_name in self.blockrefs:
            if prog.match(effective_name):
                result.extend(self.blockrefs[effective_name])

        return result

    def _select_by_type(self, type_name: str) -> List:
        filter_type = vt_int_array([0])
        filter_data = vt_variant_array([type_name])
        entities = self.select(constants.acSelectionSetAll, filter_type=filter_type, filter_data=filter_data)
        return entities

    def select_entities(self, dxf_entity: dxf.Entity) -> List:
        entities = self._select_by_type(dxf_entity.type_name)
        return [CastTo(entity, dxf_entity.interface) for entity in entities]

    def select_multi_entities(self, dxf_entities: Iterable[dxf.Entity]) -> List:
        selection = []
        for dxf_entity in set(dxf_entities):
            entities = self._select_by_type(dxf_entity.type_name)
            selection.extend([CastTo(entity, dxf_entity.interface) for entity in entities])

        return selection

    def select_all_drawing_objects(self) -> List:
        return self.select_multi_entities(dxf.AllDrawingObjects)

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

    def replace_block(self, from_block_name, to_block_name):
        self.replace_blockrefs(self.blockrefs[from_block_name], to_block_name)

    def replace_blockrefs(self, blockrefs, new_block_name):
        if not self.has_block(new_block_name):
            raise ValueError(f"There is no block named '{new_block_name}'")
        print("Start block replacing.")
        counter = 0
        for blockref in blockrefs[:]:
            self.replace_blockref(blockref, new_block_name)
            counter += 1

        print(f"Replaced {counter} blockrefs.")

    def replace_blockref(self, blockref, new_block_name):
        location = Point(*blockref.InsertionPoint)
        new_blockref = self.doc.ModelSpace.InsertBlock(
            vt_point(location),
            new_block_name,
            blockref.XScaleFactor,
            blockref.YScaleFactor,
            blockref.ZScaleFactor,
            blockref.Rotation,
            None)
        new_blockref.Layer = blockref.Layer
        copy_attributes(blockref, new_blockref)
        if new_blockref.IsDynamicBlock and blockref.IsDynamicBlock:
            copy_dynamic_properties(blockref, new_blockref)
        self.blockrefs[new_block_name].append(new_blockref)
        self.blockrefs[blockref.EffectiveName].remove(blockref)
        blockref.Delete()
        return new_blockref

    def has_block(self, name):
        for block in self.doc.Blocks:
            if block.Name == name:
                return True

        return False
