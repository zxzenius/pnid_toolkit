from utils import is_in_box
from point import Point


class Drawing:
    def __init__(self, border):
        self.border = border
        self.title_block = None
        min_point, max_point = border.GetBoundingBox()
        self.min_point = Point(*min_point)
        self.max_point = Point(*max_point)
        self.position = Point(*border.InsertionPoint)
        self.height = round(self.max_point.y - self.min_point.y)
        self.width = round(self.max_point.x - self.min_point.x)
        self.number = None
        self.row = None
        self.items = []

    @property
    def has_title(self) -> bool:
        if self.title_block:
            return True
        return False

    def __contains__(self, point: Point):
        return is_in_box(point, self.min_point, self.max_point)

    def __repr__(self):
        return f"<Drawing '{self.number}'>"

    def __str__(self):
        return str(self.number)