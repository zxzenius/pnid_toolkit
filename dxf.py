class Entity:
    def __init__(self, name: str, type_name: str = None):
        # self.type_name = type_name
        self.name = name
        if type_name is None:
            self.type_name = name.upper()
        else:
            self.type_name = type_name
        self.object_name = f"AcDb{name}"
        self.interface = f"IAcad{name}"

    def __repr__(self):
        return f"<Entity '{self.name}'>"


A3DFace = Entity("3DFace")
A3DPolyline = Entity("3DPolyline")
A3DSolid = Entity("3DSolid")
Arc = Entity("Arc")
# Attribute = Entity("Attribute", "ATTRIB")
BlockRef = Entity("BlockReference", "INSERT")
# DimAligned = Entity("DimAligned")
# DimDiametric = Entity("DimDiametric")
# DimRadialLarge = Entity("DimRadialLarge")
Circle = Entity("Circle")
Ellipse = Entity("Ellipse")
Hatch = Entity("Hatch")
Leader = Entity("Leader")
LightweightPolyline = Entity("LWPolyline")
Line = Entity("Line")
MLine = Entity("MLine")
MText = Entity("MText")
Point = Entity("Point")
Polyline = Entity("Polyline")
Region = Entity("Region")
Solid = Entity("Solid")
Spline = Entity("Spline")
Text = Entity("Text")

All2DLines = (Line, LightweightPolyline)

AllDrawingObjects = [A3DFace, A3DPolyline, A3DSolid, Arc, BlockRef, Circle, Ellipse, Hatch, Leader, LightweightPolyline,
                     Line, MLine, MText,
                     Point, Polyline, Region, Solid, Spline, Text]
