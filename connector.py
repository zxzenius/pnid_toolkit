from typing import Optional

from drawing import Drawing
from point import Point
from utils import extract_attributes, extract_dynamic_properties


class Element:
    def __init__(self, blockref):
        self.handle = blockref.Handle
        self.name = blockref.EffectiveName
        self.attrs = extract_attributes(blockref)
        self.props = extract_dynamic_properties(blockref)
        self.location = Point(*blockref.InsertionPoint)
        self.drawing: Optional[Drawing] = None
        self.ref = blockref

    def __repr__(self):
        return f"<Element '{self.name}'>"


class Connector(Element):
    cls_name = "Conn"

    @property
    def number(self) -> str:
        return self.attrs["tag"]

    @property
    def link_drawing(self) -> str:
        return self.attrs["pid.no"]

    def __repr__(self):
        return f"<{self.cls_name} '{self.handle}'>"


class UtilityConnector(Connector):
    cls_name = "UtyConn"


class MainConnector(Connector):
    cls_name = "MainConn"

    @property
    def is_entering(self) -> bool:
        left_x = self.drawing.min_point.x
        right_x = self.drawing.max_point.x
        mid_x = (left_x + right_x) / 2
        return (left_x < self.location.x < mid_x) == (not self.props["flip"])

    @property
    def is_exiting(self) -> bool:
        return not self.is_entering

    @property
    def route(self) -> str:
        return self.attrs['OriginOrDestination'.lower()]

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
        return self.props["type"] == "OFF-DRAWING"

    @property
    def is_off_boundary(self) -> bool:
        return not self.is_off_drawing
