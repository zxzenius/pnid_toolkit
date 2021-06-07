class AttributeRef:
    def __init__(self, acad_attr_ref):
        self.handle: str = acad_attr_ref.Handle
        self.tag: str = acad_attr_ref.TagString
        self.text: str = acad_attr_ref.TextString

    def __repr__(self):
        return f"<Attr '{self.tag}'='{self.text}'>"
