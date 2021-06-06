class DynamicBlockReferenceProperty:
    def __init__(self, acad_dynamic_block_reference_property):
        with acad_dynamic_block_reference_property as p:
            self.element = p

    @property
    def name(self):
        return self.element.PropertyName

    @property
    def value(self):
        return self.element.Value
