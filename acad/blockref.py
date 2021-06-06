from acad.attribute_reference import AttributeReference


class BlockRef:
    def __init__(self, acad_blockref):
        self.element = acad_blockref

    def get_attributes(self) -> dict:
        return {attr.TagString: AttributeReference(attr) for attr in self.element.GetAttributes()}

    def get_dynamic_properties(self):
        props = self.element.GetDynamicBlockProperties()
