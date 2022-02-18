from pnid import PnID
from utils import get_attribute


def get_strainers(pnid: PnID):
    return pnid.search_blockrefs(r'STRAINER_.*')


def show_strainers(pnid: PnID):
    print('===strainers===')
    counter = 1
    for strainer in get_strainers(pnid):
        if drawing := pnid.locate(strainer):
            tag = get_attribute(strainer, 'TAG').TextString
            print(f'{counter}[{drawing.tag}] {tag}')
            counter += 1


if __name__ == '__main__':
    show_strainers(PnID())
