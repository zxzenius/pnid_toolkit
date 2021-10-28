from typing import Optional

from drawing import Drawing
from point import Point
from utils import extract_attributes, extract_dynamic_properties, get_attribute, get_attributes, get_dynamic_properties


class Element:
    def __init__(self, blockref):
        self.attributes = get_attributes(blockref)
        self.dynamic_properties = get_dynamic_properties(blockref)
        self.drawing: Optional[Drawing] = None
        self.ref = blockref

    def __repr__(self):
        return f"<Element '{self.name}'>"

    def get_attribute_text(self, tag: str) -> str:
        return self.attributes[tag].TextString

    def get_dynamic_property_value(self, name: str):
        return self.dynamic_properties[name].Value

    @property
    def position(self) -> Point:
        return Point(*self.ref.InsertionPoint)

    @property
    def name(self) -> str:
        return self.ref.EffectiveName

    @property
    def handle(self) -> str:
        return self.ref.Handle


class Connector(Element):
    cls_name = "Connector"

    @property
    def tag_attr(self):
        return self.attributes["TAG"]

    @property
    def link_attr(self):
        return self.attributes["DWG.No"]

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

    def __repr__(self):
        return f"<{self.cls_name} '{self.handle}'>"


class UtilityConnector(Connector):
    cls_name = "UtyConnector"


class MainConnector(Connector):
    cls_name = "MainConnector"

    @property
    def route_attr(self):
        return self.attributes["OriginOrDestination"]

    @property
    def is_entering(self) -> bool:
        left_x = self.drawing.min_point.x
        right_x = self.drawing.max_point.x
        mid_x = (left_x + right_x) / 2
        return (left_x < self.position.x < mid_x) == (not self.dynamic_properties["flip"])

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
        return self.dynamic_properties["TYPE"].Value == "OFF-DRAWING"

    @property
    def is_off_boundary(self) -> bool:
        return not self.is_off_drawing
