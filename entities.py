# -*- coding: utf-8 -*-
import re


class Base:
    @property
    def tag(self):
        return ''

    def __repr__(self):
        return f"{self.__class__.__name__}('{str(self)}')"

    def __str__(self):
        return self.tag


# for qualified name
def gen_line_tag(service, unit, sequence, size, spec, insulation=''):
    tag = '%s%s%s-%s-%s' % (service, unit, sequence, size, spec)
    if insulation:
        tag += '-%s' % insulation
    return tag


def gen_bubble_tag(function, unit, sequence, suffix=''):
    tag = '%s-%s%s%s' % (function, unit, sequence, suffix)
    return tag


class NumberTag:
    def __init__(self, tag, unit_digits=1, has_suffix=False):
        number_match = re.fullmatch(r'(\d{%d})(\d+)([A-Z]*)' % unit_digits, tag)
        if number_match:
            self._unit, self._sequence, self._suffix = number_match.groups()
            if has_suffix == bool(self._suffix):
                return
        raise ValueError("NumberTag('%s') is not right" % tag)

    @property
    def unit(self):
        return self._unit

    @unit.setter
    def unit(self, value):
        self._unit = value

    @property
    def sequence(self):
        return self._sequence

    @sequence.setter
    def sequence(self, value):
        self._sequence = value

    @property
    def number(self):
        return self._unit + self._sequence

    @property
    def suffix(self):
        return self._suffix

    @suffix.setter
    def suffix(self, value):
        self._suffix = value


class Line(NumberTag, Base):
    def __init__(self, tag, unit_digits=2):
        """
        Line tag
        service(alpha)number(number) - size(number) - class_tag [- insulation(alpha)]
        example: NG01011-50-B2RF1-H
        :param tag:
        """
        pattern = r'([A-Z]+)(\d+)-(\d*)-(\w*)-*([A-Z]*)'
        match = re.fullmatch(pattern, tag)
        if match:
            self._service, number_tag, self._size, self._spec, self._insulation = match.groups()
            NumberTag.__init__(self, number_tag, unit_digits)
        else:
            raise ValueError('Tag "%s" is not right.' % tag)

    @property
    def tag(self):
        return gen_line_tag(self._service, self._unit, self._sequence, self._size, self._spec, self._insulation)

    @property
    def service(self):
        return self._service

    @property
    def size(self):
        return self._size

    @size.setter
    def size(self, value):
        self._size = value

    @property
    def spec(self):
        return self._spec

    @spec.setter
    def spec(self, value):
        self._spec = value

    @property
    def insulation(self):
        return self._insulation

    @insulation.setter
    def insulation(self, value):
        self._insulation = value

    @property
    def name(self):
        return self._service + self.number

    @property
    def is_insulated(self):
        return bool(self._insulation)


class Bubble(NumberTag, Base):
    def __init__(self, function, number_tag, unit_digits=2):
        func_match = re.fullmatch(r'([A-Z])([A-Z]+)', function)
        if func_match:
            NumberTag.__init__(self, number_tag, unit_digits, True)
        else:
            raise ValueError('%s %s-%s is not right' % (__class__.__name__, function, number_tag))
        self._function = function

    @property
    def tag(self):
        return gen_bubble_tag(self._function, self._unit, self._sequence, self._suffix)

    @property
    def function(self):
        return self._function

    @property
    def loop_name(self):
        return gen_bubble_tag(self.loop_type, self._unit, self._sequence)

    @property
    def loop_type(self):
        if not self.is_instrument:
            return self._function
        # 'PD', 'TD' for PDT, TDT
        elif self._function[1] == 'D':
            return self._function[:2]
        # 'PG', 'TG'
        elif self.is_gauge:
            return self._function
        else:
            return self._function[0]

    @property
    def is_valve(self):
        return self._function.endswith('V')

    @property
    def is_gauge(self):
        return self._function.endswith('G')

    @property
    def is_transmitter(self):
        return self._function.endswith('T')

    @property
    def is_sensor(self):
        return self._function.endswith('E')

    @property
    def is_instrument(self):
        return self._function not in ('PSV', 'PRV', 'FO', 'YL', 'HS', 'SC')


class Equip(Base, NumberTag):
    def __init__(self, tag, unit_digits=2):
        pattern = r'([A-Z]+)(\d+[A-Z]*)'
        match = re.fullmatch(pattern, tag)
        if match:
            self._type, number_tag = match.groups()
            self._name = None
            NumberTag.__init__(self, number_tag, unit_digits, True)
        else:
            raise ValueError('EquipTag "%s" is not right.' % tag)

    @property
    def type(self):
        return self._type

    @property
    def tag(self):
        return self.type + self.number + self.suffix

    @property
    def name(self):
        return self._name

    @name.setter
    def name(self, value):
        self._name = value


class Connector(Base, NumberTag):
    def __init__(self, tag, unit_digits=2):
        NumberTag.__init__(self, tag, unit_digits)
        self._from = None
        self._to = None

    @property
    def tag(self):
        return self.number

    @property
    def from_tag(self):
        return self._from


class Valve(Base):
    def __init__(self):
        self.handle = None
        self.pos = None
        self.v_type = None
        self.type_name = None
        self.number = None
        self.tag_handle = None

    @property
    def tag(self):
        if not self.number:
            return self.handle
        return f'{self.v_type}-{self.number}'


if __name__ == '__main__':
    print('Self Testing')
    print('Line')
    line = Line('NG010101-50-B1RF1-H', 2)
    print(repr(line))
    print('Service:%s, Unit:%s, Seq:%s, Size:%s, Spec:%s, Insulation:%s' %
          (line.service, line.unit, line.sequence, line.size, line.spec, line.insulation))
    print('Name:%s, Number:%s' % (line.name, line.number))
    print('Bubble')
    bubble = Bubble('PG', '10203B')
    print(repr(bubble))
    print(bubble.unit)
