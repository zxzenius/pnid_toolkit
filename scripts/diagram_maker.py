# Generate a diagram showing relationship of drawings and connectors
from pnid import PnID
import svgwrite


def main(pid: PnID, outfile):
    svg = svgwrite.Drawing(outfile)
    for drawing in pid.drawings:
        x = drawing.position.x
        y = drawing.position.y
        w = drawing.width
        h = drawing.height
        svg.add(svg.rect(insert=(x, y), size=(w, h), fill='gray'))
    svg.save()


def test_svg():
    svg = svgwrite.Drawing('test.svg', viewBox='-10, -10, 1000, 1000')
    svg.add(svg.rect(insert=(0, 0), size=(800, 600), fill='coral'))
    svg.save()


if __name__ == '__main__':
    # main(PnID(), 'test.svg')
    test_svg()
