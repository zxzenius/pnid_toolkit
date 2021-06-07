class DynamicProp:
    def __init__(self, acad_dynamic_prop):
        self.obj = acad_dynamic_prop

    @property
    def name(self):
        return self.obj.PropertyName

    @property
    def value(self):
        return self.obj.Value

    @value.setter
    def value(self, value):
        self.obj.Value = value

    def __repr__(self):
        return f"<Prop '{self.name}'='{self.value}'"
