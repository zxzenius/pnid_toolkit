from acad.attribute_ref import AttributeRef
from acad.dynamic_prop import DynamicProp


class BlockRef:
    def __init__(self, acad_blockref):
        self.obj = acad_blockref
        self.attrs = self.get_attrs()
        self.props = self.get_dynamic_props()

    def get_attrs(self) -> dict:
        return {attr.TagString: AttributeRef(attr) for attr in self.obj.GetAttributes()}

    def get_dynamic_props(self) -> dict:
        props = {}
        for prop in self.obj.GetDynamicBlockProperties():
            p = DynamicProp(prop)
            props[p.name] = p
        return props

    def __getattr__(self, item):
        return self.obj.item

    def __repr__(self):
        return f"<BlkRef {self.obj.Handle}>"
