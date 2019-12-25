# -*- coding: utf-8 -*-
import re
from win32com.client.gencache import EnsureDispatch
from win32com.client import CastTo
from pathlib import Path
import time
import pprint
import logging

logger = logging.getLogger('pnid')
logger.setLevel(logging.INFO)
log_handler = logging.StreamHandler()
logger.addHandler(log_handler)


def get_application(version='', visible=True):
    """
    AutoCAD Application COM Object

    :param version:
        default is latest version
        '16' for version 2006
    :param visible:
        default is True
    :return:
        AutoCAD Application Object
    """
    prog_id = 'AutoCAD.Application'
    if version:
        prog_id = '.'.join((prog_id, version))
    app = EnsureDispatch(prog_id)
    app.Visible = visible
    return app


def get_document(app, filename):
    for document in app.Documents:
        if Path(document.FullName) == Path(filename):
            return document

    return app.Documents.Open(filename)


def get_attributes(block):
    """
    Wrapper of Block.GetAttributes
    Usage: attrs = get_attributes(blockRef)
           attrs['TAG'].TextString = 'hello'
    :param block: BlockReference
    :return Dict contends AcadAttributeReference objects
    """
    attributes = dict()
    for attribute in block.GetAttributes():
        attributes[attribute.TagString.upper()] = attribute

    return attributes


# for qualified name
def gen_pipe_tag(service_tag, unit_tag, sequence, size, class_tag, insulation=''):
    pipe_tag = '%s%s%s-%s-%s' % (service_tag, unit_tag, sequence, size, class_tag)
    if insulation:
        pipe_tag += '-%s' % insulation
    return pipe_tag


def gen_instrument_tag(function_letters, unit_tag, sequence, suffix=''):
    instrument_tag = '%s-%s%s%s' % (function_letters, unit_tag, sequence, suffix)
    return instrument_tag


class Pipe:
    def __init__(self, tag):
        """
        Line tag
        service(alpha)number(number) - size(number) - class_tag [- insulation(alpha)]
        example: NG01011-50-B2RF1-H
        :param tag:
        """
        pattern = r'([A-Z]+)(\d{1})(\d+)-(\d*|DN)-(\w*)-*([A-Z]*)'
        match = re.fullmatch(pattern, tag)
        if match:
            self._service, self._unit, self._sequence, self._size, self._class_tag, self._insulation = match.groups()
            self._tag = tag
        else:
            raise ValueError('LineTag "%s" is not right.' % tag)

    def __repr__(self):
        return 'Pipe"%s"' % self._tag

    def __str__(self):
        return self._tag

    def _gen_tag(self):
        sub_tags = [self._service + self._unit + self._sequence, self._size, self._class_tag]
        if self._insulation:
            sub_tags.append(self._insulation)

        self._tag = '-'.join(sub_tags)

    @property
    def tag(self):
        return self._tag

    @property
    def service(self):
        return self._service

    @service.setter
    def service(self, value):
        if value != self._service:
            self._service = value
            self._gen_tag()

    @property
    def unit(self):
        return self._unit

    @unit.setter
    def unit(self, value):
        if value != self._unit:
            self._unit = value
            self._gen_tag()

    @property
    def sequence(self):
        return self._sequence

    @sequence.setter
    def sequence(self, value):
        if value != self._sequence:
            self._sequence = value
            self._gen_tag()

    @property
    def size(self):
        return self._size

    @size.setter
    def size(self, value):
        if value != self._size:
            self._size = value
            self._gen_tag()

    @property
    def class_tag(self):
        return self._class_tag

    @class_tag.setter
    def class_tag(self, value):
        if value != self._class_tag:
            self._class_tag = value
            self._gen_tag()

    @property
    def insulation(self):
        return self._insulation

    @insulation.setter
    def insulation(self, value):
        if value != self._insulation:
            self._insulation = value
            self._gen_tag()

    @property
    def name(self):
        return self._service + self.number

    @property
    def number(self):
        return self._unit + self._sequence


class Instrument:
    def __init__(self, function_letters, loop_tag):
        func_match = re.fullmatch(r'([A-Z])([A-Z]+)', function_letters)
        loop_match = re.fullmatch(r'(\d{1})(\d+)(.*)', loop_tag)
        if func_match and loop_match:
            self._unit, self._sequence, self._suffix = loop_match.groups()
        else:
            raise ValueError('Inst %s-%s is not right' % (function_letters, loop_tag))
        self._function = function_letters
        self._loop_tag = loop_tag

    def __repr__(self):
        return 'Inst"%s"' % self.tag

    def __str__(self):
        return self.tag

    @property
    def tag(self):
        return gen_instrument_tag(self._function, self._unit, self._sequence, self._suffix)

    @property
    def function(self):
        return self._function

    @property
    def unit(self):
        return self._unit

    @property
    def sequence(self):
        return self._sequence

    @property
    def suffix(self):
        return self._suffix

    @property
    def loop_number(self):
        return self._unit + self._sequence

    @property
    def loop_tag(self):
        return self._loop_tag

    @property
    def loop_name(self):
        return gen_instrument_tag(self.loop_letter, self._unit, self._sequence)

    @property
    def loop_letter(self):
        if self._function in ('PSV', 'PRV', 'FO'):
            return self._function
        # 'PD', 'TD' for PDT, TDT
        elif self._function[1] == 'D':
            return self._function[:2]
        # 'PG', 'TG'
        elif self.is_gauge:
            return self._function
        else:
            return self._function[0]

    @property
    def is_valve(self):
        if self._function.endswith('V'):
            return True
        else:
            return False

    @property
    def is_gauge(self):
        if self._function.endswith('G'):
            return True
        else:
            return False


class PnID:
    def __init__(self, file_path):
        self.app = get_application(version='16')
        self.doc = get_document(self.app, file_path)
        if self.doc:
            self._read()

    def _read(self):
        self.pipes = []
        self.instruments = []
        counter = 0
        self.error_pipes = []
        self.error_instruments = []

        for item in self.doc.ModelSpace:
            if item.ObjectName == 'AcDbBlockReference':
                block_ref = CastTo(item, 'IAcadBlockReference2')
                block_name = block_ref.EffectiveName.upper()
                # For pipe
                if block_name == 'PIPE_TAG':
                    self._read_pipe(get_attributes(block_ref))
                    continue
                # For Inst Bubble
                if block_name in ('DI_LOCAL', 'SH_PRI_FRONT'):
                    # attributes = get_attributes(block_ref)
                    self._read_inst(get_attributes(block_ref))
                    continue

    def _read_pipe(self, attributes):
        attr = attributes['TAG']
        try:
            self.pipes.append(Pipe(attr.TextString))
        except ValueError:
            self.error_pipes.append(attr.TextString)

    def _read_inst(self, attributes):
        function_letters = attributes['FUNCTION'].TextString
        loop_tag = attributes['TAG'].TextString
        try:
            self.instruments.append(Instrument(function_letters, loop_tag))
        except ValueError:
            self.error_instruments.append('%s-%s' % (function_letters, loop_tag))


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
    pnid = PnID(r'D:\Work\Project\XY2019P02-KAYAN.25MMSCFD.LNG\PnID\KAYAN.25.LNG.PnID_2019.1127.dwg')
    end_time = time.time()
    time_spent = end_time - start_time
    print('Time spent: %.2fs' % time_spent)

    logger.info('%s Pipes, %s Instruments' % (len(pnid.pipes), len(pnid.instruments)))
    # testing renumber 5 -> 4
    pipe_bones = gen_bones(pnid.pipes, 'service')
    inst_bones = gen_bones(pnid.instruments, 'loop_letter')
    pp.pprint(pipe_bones)
    pp.pprint(inst_bones)
    # generate mapping
    for pipe in pnid.pipes:
        new_name = '%s%02d%02d' % (pipe.service, int(pipe.unit), pipe_bones[pipe.unit][pipe.service].index(pipe.sequence)+1)
        print('%s -> %s' % (pipe.name, new_name))

    # print({inst.function for inst in pnid.instruments})
    # gen loops
    # print(gen_loops(pnid.instruments))

