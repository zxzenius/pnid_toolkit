# -*- coding: utf-8 -*-
import re
from win32com.client.gencache import EnsureDispatch
from win32com.client import CastTo
from pathlib import Path
import time
import pprint
import logging
from entities import Line, Bubble, Valve
from point import Point

logger = logging.getLogger('pnid')
logger.setLevel(logging.INFO)
log_handler = logging.StreamHandler()
logger.addHandler(log_handler)


def get_application(app='AutoCAD', version='', visible=True):
    """
    AutoCAD Application COM Object
    Also work for BricsCAD
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
    app = EnsureDispatch(prog_id)
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


def get_attributes(acad_block_ref):
    """
    Wrapper of Block.GetAttributes
    Usage: attrs = get_attributes(blockRef)
           attrs['TAG'].TextString = 'hello'
    :param acad_block_ref: BlockReference
    :return Dict contends AcadAttributeReference objects
    """
    attributes = dict()
    for attribute in acad_block_ref.GetAttributes():
        attributes[attribute.TagString.upper()] = attribute

    return attributes


def get_attribute(ent, attr_tag):
    """
    get specified attribute of ent
    :param ent: block_ref
    :param attr_tag: attribute.tagstring
    :return: AcadAttributeReference
    """
    for attribute in ent.GetAttributes():
        if attribute.TagString == attr_tag:
            return attribute

    return None


def is_in_box(point, bottom_left, top_right):
    """
    Return true if input point is in the rectangle defined by bottom_left & top_right
    :param point:
    :param bottom_left:
    :param top_right:
    :return: boolean
    """
    return (bottom_left.x < point.x < top_right.x) and (bottom_left.y < point.y < top_right.y)


class Drawing:
    def __init__(self, app_name='autocad', filepath=None):
        self.app = get_application(app=app_name)
        self.doc = get_document(self.app, filepath)

    def get_all_text(self):
        for item in self.doc.ModelSpace:
            if item.ObjectName == 'AcDbBlockReference':
                # For AutoCAD 2006, using IAcadBlockReference2
                # block_ref = CastTo(item, 'IAcadBlockReference2')
                # 2007 or higher version, using IAcadBlockReference
                block_ref = CastTo(item, 'IAcadBlockReference')
                for attribute in block_ref.GetAttributes():
                    if attribute.TextString:
                        yield attribute
            if item.ObjectName == 'AcDbMText':
                mtext = CastTo(item, 'IAcadMText')
                yield mtext
            if item.ObjectName == 'AcDbText':
                text = CastTo(item, 'IAcadText')
                yield text

    def replace(self, pattern, replacement):
        # scanning all
        regex = re.compile(pattern)
        counter = 0
        for item in self.get_all_text():
            if (result := regex.sub(replacement, item.TextString)) != item.TextString:
                item.TextString = result
                counter += 1

        print(f'{counter} items replaced.')


class PnID:
    def __init__(self, file_path):
        self.app = get_application()
        self.doc = get_document(self.app, file_path)
        if self.doc:
            self._read()
            self._post_process()

    def _read(self):
        self.lines = []
        self.bubbles = []
        self.borders = []
        self.title_blocks = dict()
        counter = 0
        self.error_lines = []
        self.error_bubbles = []
        self.valves = []

        for item in self.doc.ModelSpace:
            if item.ObjectName == 'AcDbBlockReference':
                # For AutoCAD 2006, using IAcadBlockReference2
                # block_ref = CastTo(item, 'IAcadBlockReference2')
                # 2007 or higher version, using IAcadBlockReference
                block_ref = CastTo(item, 'IAcadBlockReference')
                block_name = block_ref.EffectiveName
                # For pipe
                if block_name == 'pipe_tag':
                    # self._read_pipe(get_attributes(block_ref))
                    continue
                # For Inst Bubble
                if block_name in ('DI_LOCAL', 'SH_PRI_FRONT'):
                    # attributes = get_attributes(block_ref)
                    # self._read_inst(get_attributes(block_ref))
                    continue
                # For Border, get bottom-left & top-right coordinates of bounding box
                if block_name == 'Border.A1':
                    self._read_border(block_ref)
                    continue
                # For Title Block
                if block_name.startswith('TitleBlock.Xin'):
                    self._read_title_block(block_ref)
                    continue
                # For HandValve, not Control Valve
                if block_name.startswith('VAL_') and not block_name.startswith('VAL_CTRL'):
                    self._read_valve(block_ref)
                    continue

    def _read_pipe(self, attributes):
        attr = attributes['TAG']
        try:
            self.lines.append(Line(attr.TextString))
        except ValueError:
            self.error_lines.append(attr.TextString)

    def _read_inst(self, attributes):
        function_letters = attributes['FUNCTION'].TextString
        loop_tag = attributes['TAG'].TextString
        try:
            self.bubbles.append(Bubble(function_letters, loop_tag))
        except ValueError:
            self.error_bubbles.append('%s-%s' % (function_letters, loop_tag))

    def _read_border(self, block_ref):
        """
        Get btm-left & top-right of border block using GetBoundingBox, append to borders
        :param block_ref:
        :return:
        """
        bottom_left, top_right = block_ref.GetBoundingBox()
        self.borders.append((Point(bottom_left), Point(top_right)))

    def _read_title_block(self, block_ref):
        """
        Get dwg_no & insertion point of title block, insert to a dict to avoid duplicated dwg_no
        :param block_ref:
        :return:
        """
        dwg_number = get_attribute(block_ref, 'DWG.NO.').TextString
        if dwg_number:
            self.title_blocks[dwg_number] = Point(block_ref.InsertionPoint)

    def _read_valve(self, block_ref):
        valve = Valve()
        valve.handle = block_ref.Handle
        valve.pos = Point(block_ref.InsertionPoint)
        valve.tag_handle = get_attribute(block_ref, 'TAG').Handle
        valve.type_name = block_ref.EffectiveName[4::]
        self.valves.append(valve)

    def gen_dwg_map(self):
        """
        Generate dwg_no <-> coordinates for entity locating
        :return: a dict
        """
        result = dict()
        for dwg_number, insert_point in self.title_blocks.items():
            for bottom_left, top_right in self.borders:
                if is_in_box(insert_point, bottom_left, top_right):
                    result[dwg_number] = (bottom_left, top_right)
                    break
        return result

    def locate_dwg_no(self, point):
        """
        Locate drawing number for input point
        :param point: Point
        :return: drawing number
        """
        for dwg_no, (bottom_left, top_right) in self.dwg_map.items():
            if is_in_box(point, bottom_left, top_right):
                return dwg_no
        return None

    def _post_process(self):
        self.dwg_map = self.gen_dwg_map()


def gen_loops(instruments):
    loops = dict()
    for instrument in instruments:
        loop_name = instrument.loop_name
        if loop_name not in loops:
            loops[loop_name] = []
        loops[loop_name].append(str(instrument))

    return loops


def gen_bones(items, keyword_attribute):
    bones = dict()
    # collecting
    for item in items:
        unit = item.unit
        keyword = item.__getattribute__(keyword_attribute)
        if unit not in bones:
            bones[unit] = dict()
        if keyword not in bones[unit]:
            bones[unit][keyword] = set()
        bones[unit][keyword].add(item.sequence)
    # arranging
    for bone_key in bones.keys():
        for keyword in bones[bone_key].keys():
            bones[bone_key][keyword] = sorted(bones[bone_key][keyword])

    return bones


if __name__ == '__main__':
    # line = Line('NG01010-50-B2SRF1-H')
    # print(line)
    pp = pprint.PrettyPrinter()
    start_time = time.time()
    pnid = PnID(r'D:\Work\Project\XY2019P02-KAYAN.25MMSCFD.LNG\PnID\IFC\KAYAN.LNG.TRAIN.ColdBox.PnID_2020.0605.dwg')
    end_time = time.time()
    time_spent = end_time - start_time
    print('Time spent: %.2fs' % time_spent)

    # logger.info('%s Pipes, %s Instruments' % (len(pnid.lines), len(pnid.bubbles)))
    # testing renumber 5 -> 4
    # pipe_bones = gen_bones(pnid.lines, 'service')
    # inst_bones = gen_bones(pnid.bubbles, 'loop_letter')
    # pp.pprint(pipe_bones)
    # pp.pprint(inst_bones)
    # generate mapping
    # for pipe in pnid.lines:
    #     new_name = '%s%02d%02d' % (pipe.service, int(pipe.unit), pipe_bones[pipe.unit][pipe.service].index(pipe.sequence)+1)
    #     print('%s -> %s' % (pipe.name, new_name))

    # print({inst.function for inst in pnid.instruments})
    # gen loops
    # print(gen_loops(pnid.instruments))
    # pp.pprint(pnid.borders)
    # pp.pprint(pnid.title_blocks)
    # pp.pprint(pnid.locate_dwg_no(Point((298, -3590, 0))))
    # pp.pprint(pnid.dwg_map)
    # pp.pprint(pnid.valves)
    print(f'Acad.app.version: {pnid.app.version}')
