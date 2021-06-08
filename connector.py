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

    def __repr__(self):
        return f"<Element '{self.name}'>"


class Connector(Element):
    @property
    def number(self) -> str:
        return self.attrs["tag"]


class UtilityConnector(Connector):
    def __repr__(self):
        return f"<UtyConn '{self.number}'>"


class MainConnector(Connector):
    def __repr__(self):
        return f"<MainConn '{self.number}'>"

    @property
    def is_facing_in(self) -> bool:
        left_x = self.drawing.min_point.x
        right_x = self.drawing.max_point.x
        mid_x = (left_x + right_x) / 2
        return (left_x < self.location.x < mid_x) & (not self.props["flip"])

    @property
    def is_facing_out(self) -> bool:
        return not self.is_facing_in

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

