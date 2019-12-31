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
file_handler = logging.FileHandler(Path('change.log'), encoding='utf8')
logger.addHandler(log_handler)
logger.addHandler(file_handler)


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
        self.category = 'Pipe'

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
        self.category = 'Instrument'

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

    @unit.setter
    def unit(self, value):
        self._unit = value

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
        return self.loop_number + self._suffix

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


class Inline:
    def __init__(self, tag):
        pattern = r'([A-Z]+)-(\d{1})(\d+)([A-Z]*)'
        match = re.fullmatch(pattern, tag)
        if match:
            self._category, self._unit, self._sequence, self._suffix = match.groups()
        else:
            raise ValueError('InLine "%s" is not right.' % tag)
        self.category = self._category

    @property
    def unit(self):
        return self._unit

    @unit.setter
    def unit(self, value):
        self._unit = value

    @property
    def tag(self):
        return '%s-%s%s%s' % (self._category, self._unit, self._sequence, self._suffix)

    def __repr__(self):
        return 'Inline"%s"' % self.tag

    def __str__(self):
        return self.tag


class Connector:
    def __init__(self, tag):
        pattern = r'(\d{1})(\d+)'
        match = re.fullmatch(pattern, tag)
        if match:
            self._unit, self._sequence = match.groups()
        else:
            raise ValueError('Connector "%s" is not right.' % tag)
        self.category = 'Connector'

    @property
    def unit(self):
        return self._unit

    @unit.setter
    def unit(self, value):
        self._unit = value

    @property
    def tag(self):
        return self._unit + self._sequence


class PnID:
    def __init__(self, file_path):
        self.app = get_application(version='16')
        self.doc = get_document(self.app, file_path)
        if self.doc:
            self._renumber()

    def _renumber(self):
        self.pipes = []
        self.instruments = []
        self.inlines = []
        self.connectors = []
        self.text_counter = 0
        counter = 0
        self.error_pipes = []
        self.error_instruments = []

        for item in self.doc.ModelSpace:
            if item.ObjectName == 'AcDbBlockReference':
                block_ref = CastTo(item, 'IAcadBlockReference2')
                block_name = block_ref.EffectiveName.upper()
                # For pipe
                if block_name == 'PIPE_TAG':
                    self._process_pipe(get_attributes(block_ref))
                    continue
                # For Inst Bubble
                if block_name in ('DI_LOCAL', 'SH_PRI_FRONT', 'SC_LOCAL'):
                    self._process_inst(get_attributes(block_ref))
                    continue
                # For Strainer
                if block_name.startswith('STRAINER'):
                    self._process_inline(get_attributes(block_ref))
                    continue
                # For Connector
                if block_name in ('CONNECTOR_MAIN', 'CONNECTOR_UTILITY'):
                    self._process_connector(get_attributes(block_ref))
                    continue
            elif item.ObjectName == 'AcDbText':
                text = CastTo(item, 'IAcadText')
                match = re.fullmatch(r'.*-(\d{5}).*', text.TextString)
                if match:
                    old_tag = match.group(1)
                    new_tag = '0' + old_tag
                    logger.info('Text: %s -> %s' % (text.TextString, new_tag))
                    self.text_counter += 1
                    text.TextString = text.TextString.replace(old_tag, new_tag)

    def _process_pipe(self, attributes):
        attr = attributes['TAG']
        try:
            pipe = Pipe(attr.TextString)
            self.pipes.append(pipe)
            tag_change(pipe)
            attr.TextString = pipe.tag
        except ValueError:
            self.error_pipes.append(attr.TextString)

    def _process_inst(self, attributes):
        function_letters = attributes['FUNCTION'].TextString
        loop_tag = attributes['TAG'].TextString
        try:
            instrument = Instrument(function_letters, loop_tag)
            self.instruments.append(instrument)
            tag_change(instrument)
            attributes['TAG'].TextString = instrument.loop_tag
            # attributes['TAG'].ScaleFactor = 0.6
        except ValueError:
            self.error_instruments.append('%s-%s' % (function_letters, loop_tag))

    def _process_inline(self, attributes):
        attr = attributes['TAG']
        try:
            inline = Inline(attr.TextString)
            self.inlines.append(inline)
            tag_change(inline)
            attr.TextString = inline.tag
        except ValueError:
            # self.error_pipes.append(attr.TextString)
            pass

    def _process_connector(self, attributes):
        attr = attributes['TAG']
        try:
            connector = Connector(attr.TextString)
            self.connectors.append(connector)
            tag_change(connector)
            attr.TextString = connector.tag
            # attr.ScaleFactor = 0.7
            if 'ORIGINORDESTINATION' in attributes:
                from_to = attributes['ORIGINORDESTINATION']
                match = re.fullmatch(r'(FROM|TO).*(\d{1})(\d{4}).*', from_to.TextString)
                if match:
                    temp_str, unit, seq = match.groups()
                    old_number = '%s%s' % (unit, seq)
                    new_number = '0%s%s' % (unit, seq)
                    logger.info('%s -> %s' % (from_to.TextString, new_number))
                    from_to.TextString = from_to.TextString.replace(old_number, new_number)

        except ValueError:
            pass


def tag_change(element):
    origin_name = element.tag
    element.unit = '0' + element.unit
    new_name = element.tag
    logger.info('%s: %s -> %s' % (element.category, origin_name, new_name))


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
    target_file = r'D:\Work\Project\XY2019P02-KAYAN.25MMSCFD.LNG\PnID\Loading\KAYAN.25.Loading.PnID_2019.1226.dwg'
    logger.info('Start with "%s"' % target_file)
    pnid = PnID(target_file)
    end_time = time.time()
    time_spent = end_time - start_time
    print('Time spent: %.2fs' % time_spent)

    logger.info('%s Pipes, %s Instruments, %s Inlines, %s Connectors, %s Texts' % (len(pnid.pipes), len(pnid.instruments), len(pnid.inlines), len(pnid.connectors), pnid.text_counter))
