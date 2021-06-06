class AttributeReference:
    def __init__(self, acad_attribute_reference):
        self.handle: str = acad_attribute_reference.Handle
        self.tag: str = acad_attribute_reference.TagString
        self.text: str = acad_attribute_reference.TextString

    def __repr__(self):
        return f"<Attr '{self.tag}'='{self.text}'>"
