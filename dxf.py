class Entity:
    def __init__(self, type_name: str, name: str):
        self.type_name = type_name
        self.name = name
        self.object_name = f"AcDb{name}"
        self.interface = f"IAcad{name}"

    def __repr__(self):
        return f"<Entity '{self.name}'>"


Arc = Entity("ARC", "Arc")
BlockRef = Entity("INSERT", "BlockReference")
Circle = Entity("CIRCLE", "Circle")
Ellipse = Entity("ELLIPSE", "Ellipse")
Line = Entity("LINE", "Line")
MLine = Entity("MLINE", "MLine")
MText = Entity("MTEXT", "MText")
Point = Entity("POINT", "Point")
Polyline = Entity("POLYLINE", "PolyLine")
Spline = Entity("SPLINE", "Spline")
Text = Entity("TEXT", "Text")
