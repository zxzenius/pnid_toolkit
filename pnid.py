from collections import defaultdict
from typing import List

from caddoc import CADDoc
from drawing import Drawing
from components import MainConnector, UtilityConnector, Bubble, Line
from point import Point


def gen_dwg_no(unit: int, seq: int) -> str:
    return f"{unit:02d}{seq:02d}"


def tagging_with_unit(drawings: List[Drawing], start_unit: int = 1, start_seq: int = 1):
    unit = start_unit - 1
    seq = start_seq
    index = None
    for drawing in drawings:
        if drawing.has_title:
            if index != drawing.row:
                seq = 1
                unit += 1
                index = drawing.row
            else:
                seq += 1
            drawing._number = gen_dwg_no(unit, seq)


def sorted_drawings(drawings: List[Drawing]):
    min_height = min((h.height for h in drawings))
    sorted_by_x = sorted(drawings, key=lambda dwg: dwg.position.x)
    db = defaultdict(list)
    for drawing in sorted_by_x:
        y = drawing.position.y
        index = round(y / min_height)
        db[index].append(drawing)
    counter = 1
    result = []
    for key in sorted(db.keys(), reverse=True):
        for drawing in db[key]:
            # row tagging
            drawing.row = counter
            result.append(drawing)
        counter += 1
    return result


# todo: outline (mark) target entity for easy searching manually
class PnID(CADDoc):
    def __init__(self, filepath: str = None):
        self.drawings: List[Drawing] = []
        self.main_connectors: List[MainConnector] = []
        self.utility_connectors: List[UtilityConnector] = []
        self.bubbles: List[Bubble] = []
        self.lines: List[Line] = []
        super().__init__(filepath=filepath)

    def init_db(self):
        super().init_db()
        self.load_drawings()
        self.load_connectors()
        self.load_bubbles()
        self.load_lines()

    def get_title_blocks(self):
        return self.search_blockrefs("^TitleBlock.*")

    def get_borders(self):
        return self.search_blockrefs("^Border.*")

    def load_drawings(self):
        print("Loading drawings")
        borders = self.get_borders()
        title_blocks = self.get_title_blocks()
        drawings = []
        for border in borders:
            drawing = Drawing(border)
            drawings.append(drawing)
            for title_block in title_blocks[:]:
                if Point(*title_block.InsertionPoint) in drawing:
                    drawing.title_block = title_block
                    title_blocks.remove(title_block)
                    break

        self.drawings = drawings
        self.sort_drawings()

    def sort_drawings(self):
        self.drawings = sorted_drawings(self.drawings)

    def load_connectors(self):
        print("Loading connectors")
        self.main_connectors = self.get_main_connectors()
        print(f"{len(self.main_connectors)} main connectors.")
        self.utility_connectors = self.get_utility_connectors()
        print(f"{len(self.utility_connectors)} utility connectors.")

    def get_main_connectors(self):
        return self.wrap_blockrefs(self.blockrefs.get('Connector_Main', []), MainConnector)

    def get_utility_connectors(self):
        return self.wrap_blockrefs(self.blockrefs.get('Connector_Utility', []), UtilityConnector)

    def load_bubbles(self):
        print('Loading bubbles')
        self.bubbles = self.get_bubbles()
        print(f'{len(self.bubbles)} bubbles.')

    def get_bubbles(self) -> List[Bubble]:
        bubbles = self.search_blockrefs(r'\w+_(LOCAL|FRONT|BACK)')
        return self.wrap_blockrefs(bubbles, Bubble)

    def load_lines(self):
        print('Loading Lines')
        self.lines = self.get_lines()
        print(f'{len(self.lines)} lines.')

    def get_lines(self) -> List[Line]:
        lines = self.blockrefs.get('pipe_tag', []) + self.blockrefs.get('TAG_NUMBER', [])
        return self.wrap_blockrefs(lines, Line)

    def wrap_blockrefs(self, blockrefs: List, wrapper):
        return [self.wrap_blockref(blockref, wrapper) for blockref in blockrefs]

    def wrap_blockref(self, blockref, wrapper):
        target = wrapper(blockref)
        target.drawing = self.locate(blockref)
        return target

    def locate(self, blockref):
        for drawing in self.drawings:
            if Point(*blockref.InsertionPoint) in drawing:
                return drawing
        return None


def gen_loops(instruments):
    loops = dict()
    for instrument in instruments:
        loop_name = instrument.loop_name
        if loop_name not in loops:
            loops[loop_name] = []
        loops[loop_name].append(str(instrument))

    return loops


def gen_bones(items, keyword_attribute):
    bones = dict()
    # collecting
    for item in items:
        unit = item.unit
        keyword = item.__getattribute__(keyword_attribute)
        if unit not in bones:
            bones[unit] = dict()
        if keyword not in bones[unit]:
            bones[unit][keyword] = set()
        bones[unit][keyword].add(item.sequence)
    # arranging
    for bone_key in bones.keys():
        for keyword in bones[bone_key].keys():
            bones[bone_key][keyword] = sorted(bones[bone_key][keyword])

    return bones


if __name__ == '__main__':
    pnid = PnID()
