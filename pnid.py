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
        # self.logger = logging.getLogger(__name__)
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

    # def get_local_bubbles(self) -> dict:
    #     """
    #     Index all local bubbles
    #     :return: db[code][tag](blockref)
    #     """
    #     print('Loading bubbles')
    #     db = {}
    #     for bubble in self.search_blockrefs(r'\w*_LOCAL$'):
    #         code = get_attribute(bubble, 'FUNCTION').TextString
    #         number = get_attribute(bubble, 'TAG').TextString
    #         tag = f'{code}-{number}'
    #         if code not in db.keys():
    #             db[code] = defaultdict(list)
    #         db[code][tag].append(bubble)
    #     return db

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
    #
    # def _read(self):
    #     self.lines = []
    #     self.bubbles = []
    #     self.borders = []
    #     self.title_blocks = dict()
    #     counter = 0
    #     self.error_lines = []
    #     self.error_bubbles = []
    #     self.valves = []
    #
    #     for item in self.doc.ModelSpace:
    #         if item.ObjectName == 'AcDbBlockReference':
    #             # For AutoCAD 2006, using IAcadBlockReference2
    #             # block_ref = CastTo(item, 'IAcadBlockReference2')
    #             # 2007 or higher version, using IAcadBlockReference
    #             block_ref = CastTo(item, 'IAcadBlockReference')
    #             block_name = block_ref.EffectiveName
    #             # For pipe
    #             if block_name == 'pipe_tag':
    #                 # self._read_pipe(get_attributes(block_ref))
    #                 continue
    #             # For Inst Bubble
    #             if block_name in ('DI_LOCAL', 'SH_PRI_FRONT'):
    #                 # attributes = get_attributes(block_ref)
    #                 # self._read_inst(get_attributes(block_ref))
    #                 continue
    #             # For Border, get bottom-left & top-right coordinates of bounding box
    #             if block_name == 'Border.A1':
    #                 self._read_border(block_ref)
    #                 continue
    #             # For Title Block
    #             if block_name.startswith('TitleBlock.Xin'):
    #                 self._read_title_block(block_ref)
    #                 continue
    #             # For HandValve, not Control Valve
    #             if block_name.startswith('VAL_') and not block_name.startswith('VAL_CTRL'):
    #                 self._read_valve(block_ref)
    #                 continue
    #
    # def _read_pipe(self, attributes):
    #     attr = attributes['TAG']
    #     try:
    #         self.lines.append(Line(attr.TextString))
    #     except ValueError:
    #         self.error_lines.append(attr.TextString)
    #
    # def _read_inst(self, attributes):
    #     function_letters = attributes['FUNCTION'].TextString
    #     loop_tag = attributes['TAG'].TextString
    #     try:
    #         self.bubbles.append(Bubble(function_letters, loop_tag))
    #     except ValueError:
    #         self.error_bubbles.append('%s-%s' % (function_letters, loop_tag))
    #
    # def _read_border(self, block_ref):
    #     """
    #     Get btm-left & top-right of border block using GetBoundingBox, append to borders
    #     :param block_ref:
    #     :return:
    #     """
    #     bottom_left, top_right = block_ref.GetBoundingBox()
    #     self.borders.append((Point(bottom_left), Point(top_right)))
    #
    # def _read_title_block(self, block_ref):
    #     """
    #     Get dwg_no & insertion point of title block, insert to a dict to avoid duplicated dwg_no
    #     :param block_ref:
    #     :return:
    #     """
    #     dwg_number = get_attribute(block_ref, 'DWG.NO.').TextString
    #     if dwg_number:
    #         self.title_blocks[dwg_number] = Point(block_ref.InsertionPoint)
    #
    # def _read_valve(self, block_ref):
    #     valve = Valve()
    #     valve.handle = block_ref.Handle
    #     valve.pos = Point(*block_ref.InsertionPoint)
    #     valve.tag_handle = get_attribute(block_ref, 'TAG').Handle
    #     valve.type_name = block_ref.EffectiveName[4::]
    #     self.valves.append(valve)

    # def read(self):
    #     self.lines = []
    #     self.bubbles = []
    #     self.borders = []
    #     self.title_blocks = dict()
    #     counter = 0
    #     self.error_lines = []
    #     self.error_bubbles = []
    #     self.valves = []
    #
    #     for item in self.doc.ModelSpace:
    #         if item.ObjectName == 'AcDbBlockReference':
    #             # For AutoCAD 2006, using IAcadBlockReference2
    #             # block_ref = CastTo(item, 'IAcadBlockReference2')
    #             # 2007 or higher version, using IAcadBlockReference
    #             block_ref = CastTo(item, 'IAcadBlockReference')
    #             block_name = block_ref.EffectiveName
    #             # For pipe
    #             if block_name == 'pipe_tag':
    #                 # self._read_pipe(get_attributes(block_ref))
    #                 continue
    #             # For Inst Bubble
    #             if block_name in ('DI_LOCAL', 'SH_PRI_FRONT'):
    #                 # attributes = get_attributes(block_ref)
    #                 # self._read_inst(get_attributes(block_ref))
    #                 continue
    #             # For Border, get bottom-left & top-right coordinates of bounding box
    #             if block_name == 'Border.A1':
    #                 self._read_border(block_ref)
    #                 continue
    #             # For Title Block
    #             if block_name.startswith('TitleBlock.Xin'):
    #                 self._read_title_block(block_ref)
    #                 continue
    #             # For HandValve, not Control Valve
    #             if block_name.startswith('VAL_') and not block_name.startswith('VAL_CTRL'):
    #                 self._read_valve(block_ref)
    #                 continue
    #
    # def gen_dwg_map(self):
    #     """
    #     Generate dwg_no <-> coordinates for entity locating
    #     :return: a dict
    #     """
    #     result = dict()
    #     for dwg_number, insert_point in self.title_blocks.items():
    #         for bottom_left, top_right in self.borders:
    #             if is_in_box(insert_point, bottom_left, top_right):
    #                 result[dwg_number] = (bottom_left, top_right)
    #                 break
    #     return result
    #
    # def locate_dwg_no(self, point):
    #     """
    #     Locate drawing number for input point
    #     :param point: Point
    #     :return: drawing number
    #     """
    #     for dwg_no, (bottom_left, top_right) in self.dwg_map.items():
    #         if is_in_box(point, bottom_left, top_right):
    #             return dwg_no
    #     return None
    #
    # def _post_process(self):
    #     self.dwg_map = self.gen_dwg_map()


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
    pid = PnID()
    # end_time = time.time()
    # time_spent = end_time - start_time
    # print('Time spent: %.2fs' % time_spent)
