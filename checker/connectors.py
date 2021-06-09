from connector import Connector
from pnid import PnID


def problem_line(conn: Connector, problem: str) -> dict:
    return {
        "problem": problem,
        "target": conn.handle,
        "number": conn.number,
        "drawing": conn.drawing.number,
        "location": (conn.location.x, conn.location.y)
    }


def number_matched(conn: Connector) -> bool:
    return conn.number[:3] == conn.drawing.number[1:]


def check(p: PnID):
    connectors = p.main_connectors
    for connector in connectors:
        if not connector.number:
            print(problem_line(connector, "Missing number"))
        elif not (connector.is_to | connector.is_from):
            print(problem_line(connector, "Missing route"))
        elif connector.is_entering != connector.is_from:
            print(problem_line(connector, "Wrong direction"))
        elif connector.is_to & (not number_matched(connector)):
            print(problem_line(connector, "Wrong number when exiting"))
        elif connector.is_from & number_matched(connector):
            print(problem_line(connector, "Wrong number when entering"))
        elif connector.is_off_boundary & bool(connector.link_drawing):
            print(problem_line(connector, "P&ID No. not blank in off-boundary conn"))
        elif connector.is_off_drawing & (not bool(connector.link_drawing)):
            print(problem_line(connector, "Missing P&ID No. in off-drawing conn"))
