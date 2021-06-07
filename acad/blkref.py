from acad.attribute_ref import AttributeRef
from acad.dynamic_prop import DynamicProp


class BlkRef:
    def __init__(self, acad_blockref):
        self.obj = acad_blockref
        self.attrs = self.get_attrs()
        self.props = self.get_dynamic_props()
        self.drawing = None

    def get_attrs(self) -> dict:
        return {attr.TagString: attr for attr in self.obj.GetAttributes()}

    def get_dynamic_props(self) -> dict:
        return {prop.PropertyName: prop for prop in self.obj.GetDynamicBlockProperties()}

    # def __getattr__(self, item):
    #     return self.obj.item

    def __repr__(self):
        return f"<BlkRef {self.obj.Handle}>"
