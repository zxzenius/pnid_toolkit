from typing import Union

from utils import is_in_box, extract_attributes, get_attributes
from point import Point


class Drawing:
    def __init__(self, border):
        self.border = border
        self._title_block = None
        min_point, max_point = border.GetBoundingBox()
        self.min_point = Point(*min_point)
        self.max_point = Point(*max_point)
        self.position = Point(*border.InsertionPoint)
        self.height = round(self.max_point.y - self.min_point.y)
        self.width = round(self.max_point.x - self.min_point.x)
        self._number = None
        self.row = None
        self.items = []
        self._attrs = {}

    @property
    def has_title(self) -> bool:
        if self._title_block:
            return True
        return False

    def __contains__(self, point: Point):
        return is_in_box(point, self.min_point, self.max_point)

    def __repr__(self):
        return f"<Drawing '{self.tag}'>"

    def __str__(self):
        return str(self._number)

    @property
    def title_block(self):
        return self._title_block

    @title_block.setter
    def title_block(self, blockref):
        self._title_block = blockref
        self._attrs = get_attributes(blockref)

    @property
    def id(self) -> Union[None, str]:
        if not self.title_block:
            return None
        return self._attrs["DWG.NO."].TextString

    @id.setter
    def id(self, value: str):
        if self.title_block:
            self._attrs["DWG.NO."].TextString = value

    @property
    def tag(self):
        if self.id:
            return self.id[-4:]

    @property
    def number(self):
        return self._number

    @number.setter
    def number(self, value: str):
        self._number = value
