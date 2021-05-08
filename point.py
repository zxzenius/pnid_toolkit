# -*- coding: utf-8 -*-
class Point:
    def __init__(self, point_tuple):
        self.x, self.y, self.z = point_tuple

    def __repr__(self):
        return f'Point({self.x}, {self.y}, {self.z})'
