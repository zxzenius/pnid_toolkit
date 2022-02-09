import re
from collections import namedtuple
from typing import Optional, NamedTuple

from drawing import Drawing
from point import Point
from utils import extract_attributes, extract_dynamic_properties, get_attribute, get_attributes, get_dynamic_properties


class BlockRefWrapper:
    """
    BlockRef Wrapper
    """
    def __init__(self, blockref):
        self.attributes = get_attributes(blockref)
        self.dynamic_properties = get_dynamic_properties(blockref)
        self.drawing: Optional[Drawing] = None
        self.ent = blockref

    def __repr__(self):
        return f"BlockRef('{self.name}')"

    def get_attribute_text(self, tag: str) -> str:
        return self.attributes[tag].TextString

    def get_dynamic_property_value(self, name: str):
        return self.dynamic_properties[name].Value

    @property
    def position(self) -> Point:
        return Point(*self.ent.InsertionPoint)

    @property
    def name(self) -> str:
        return self.ent.EffectiveName

    @property
    def handle(self) -> str:
        return self.ent.Handle


class Component(BlockRefWrapper):
    def __repr__(self):
        if self.tag:
            tag = f"'{self.tag}'"
        else:
            tag = 'Untagged'

        return f"{self.__class__.__name__}({tag})"

    @property
    def tag(self):
        return ''


class Connector(Component):
    @property
    def tag_attr(self):
        return self.attributes["TAG"]

    @property
    def link_attr(self):
        return self.attributes["PID.No"]

    @property
    def service_attr(self):
        return self.attributes["Service"]

    @property
    def description_attr(self):
        return self.attributes["DESC"]

    @property
    def tag(self) -> str:
        return self.tag_attr.TextString

    @property
    def link_drawing(self) -> str:
        return self.link_attr.TextString


class UtilityConnector(Connector):
    pass


class MainConnector(Connector):
    @property
    def route_attr(self):
        return self.attributes["OriginOrDestination"]

    @property
    def is_flip(self) -> bool:
        return self.get_dynamic_property_value("Flip")

    @property
    def is_entering(self) -> bool:
        left_x = self.drawing.min_point.x
        right_x = self.drawing.max_point.x
        mid_x = (left_x + right_x) / 2
        return (left_x < self.position.x < mid_x) == (not self.is_flip)

    @property
    def is_exiting(self) -> bool:
        return not self.is_entering

    @property
    def route(self) -> str:
        return self.route_attr.TextString

    @route.setter
    def route(self, value: str):
        self.route_attr.TextString = value

    @property
    def is_to(self) -> bool:
        for keyword in ('TO', '至'):
            if self.route.startswith(keyword):
                return True
        return False

    @property
    def is_from(self) -> bool:
        for keyword in ('FROM', '自'):
            if self.route.startswith(keyword):
                return True
        return False

    @property
    def is_off_drawing(self) -> bool:
        return self.get_dynamic_property_value("TYPE") == "OFF-DRAWING"

    @property
    def is_off_boundary(self) -> bool:
        return not self.is_off_drawing


class Bubble(Component):
    @property
    def code(self):
        return self.get_attribute_text('FUNCTION')

    @property
    def number(self):
        return self.get_attribute_text('TAG')

    @property
    def is_gauge(self):
        return self.code.endswith('G')

    @property
    def is_transmitter(self):
        return self.code.endswith('T')

    @property
    def is_sensor(self):
        return self.code.endswith('E')

    @property
    def is_valve(self):
        return self.code.endswith('V')

    @property
    def is_instrument(self):
        return self.code not in ['PSV', 'PRV', 'FO', 'YL', 'HS', 'SC']

    @property
    def tag(self):
        return f'{self.code}-{self.number}'

    @property
    def loop_code(self):
        if not self.is_instrument:
            return self.code
            # 'PD', 'TD' for PDT, TDT
        elif self.code[1] == 'D':
            return self.code[:2]
        # 'PG', 'TG'
        elif self.is_gauge:
            return self.code
        else:
            return self.code[0]


class Line(Component):
    def __init__(self, blockref):
        super().__init__(blockref)
        self._service, self._number, self._size, self._spec, self._insulation = parse_line_tag(self.raw_tag)

    def get_tag(self):
        return self.get_attribute_text('TAG')

    def gen_tag(self):
        return gen_line_tag(self.service, self.number, self.size, self.spec, self.insulation)

    @property
    def raw_tag(self):
        """
        'TAG' prop of blockref
        :return:
        """
        return self.get_tag()

    @property
    def tag(self):
        """
        Generated tag
        :return:
        """
        return self.gen_tag()

    def sync(self):
        """
        Sync 'TAG' of blockref with generated tag
        :return:
        """
        self.attributes['TAG'] = self.tag

    @property
    def service(self):
        return self._service

    @service.setter
    def service(self, value: str):
        self._service = value
        self.sync()

    @property
    def number(self):
        return self._number

    @number.setter
    def number(self, value: str):
        self._number = value
        self.sync()

    @property
    def size(self):
        return self._size

    @size.setter
    def size(self, value: str):
        self._size = value
        self.sync()

    @property
    def spec(self):
        return self._spec

    @spec.setter
    def spec(self, value: str):
        self._spec = value
        self.sync()

    @property
    def insulation(self):
        return self._insulation

    @insulation.setter
    def insulation(self, value: str):
        self._insulation = value
        self.sync()


class LineTag(NamedTuple):
    service: str
    number: str
    size: str
    spec: str
    insulation: str


def parse_line_tag(tag: str):
    pattern = r'([A-Z]+|\?)(\d+|\?)-(\w*|\?)-(\w*|\?)-*([A-Z]*)'
    match = re.fullmatch(pattern, tag)
    service = ''
    number = ''
    size = ''
    spec = ''
    insulation = ''
    if match:
        service, number, size, spec, insulation = match.groups()
    return LineTag(service, number, size, spec, insulation)


def gen_line_tag(service: str, number: str, size: str, spec: str, insulation: str):
    tag = f'{placeholder(service)}{placeholder(number)}-{placeholder(size)}-{placeholder(spec)}'
    if insulation:
        tag = f'{tag}-{insulation}'
    return tag


def placeholder(text: str):
    if not text:
        text = '?'
    return text


class PureLine:
    pass
