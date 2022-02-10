from config import load_config
from pnid import PnID


class Checker:
    def __init__(self, pnid: PnID, conf=None):
        self.pnid = pnid
        self.config = load_config(conf)

    def check(self):
        pass

    def check_connectors(self):
        pass

    def make_report(self):
        pass


if __name__ == '__main__':
    Checker(pnid=PnID()).check()
