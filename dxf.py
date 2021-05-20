from typing import NamedTuple


class Entity(NamedTuple):
    type_name: str
    interface: str


Arc = Entity("ARC", "IAcadArc")
BlockRef = Entity("INSERT", "IAcadBlockReference")
Circle = Entity("CIRCLE", "IAcadCircle")
Ellipse = Entity("ELLIPSE", "IAcadEllipse")
Line = Entity("LINE", "IAcadLine")
MLine = Entity("MLINE", "IAcadMLine")
MText = Entity("MTEXT", "IAcadMText")
Point = Entity("POINT", "IAcadPoint")
Polyline = Entity("POLYLINE", "IAcadPolyLine")
Spline = Entity("SPLINE", "IAcadSpline")
Text = Entity("TEXT", "IAcadText")
