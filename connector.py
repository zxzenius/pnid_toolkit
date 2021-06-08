from typing import Optional

from drawing import Drawing
from point import Point
from utils import extract_attributes, extract_dynamic_properties


class Component:
    def __init__(self, blockref):
        self.handle = blockref.Handle
        self.name = blockref.EffectiveName
        self.attrs = extract_attributes(blockref)
        self.props = extract_dynamic_properties(blockref)
        self.location = Point(*blockref.InsertionPoint)
        self.drawing: Optional[Drawing] = None

    def __repr__(self):
        return f"<Component>"


class MainConnector(Component):
    def __init__(self, blockref):
        if blockref.EffectiveName != "Connector_Main":
            raise TypeError(f"{blockref.EffectiveName} is not Connector")
        super().__init__(blockref)

    @property
    def number(self) -> str:
        return self.attrs["tag"]

    @property
    def is_income(self) -> bool:
        left_x = self.drawing.min_point.x
        right_x = self.drawing.max_point.x
        mid_x = (left_x + right_x) / 2
        return (left_x < self.location.x < mid_x) & (not self.props["flip"])

    @property
    def is_outcome(self) -> bool:
        return not self.is_income
