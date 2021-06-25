from connector import Connector
from pnid import PnID
from pprint import PrettyPrinter


def problem_line(conn: Connector, problem: str) -> dict:
    return {
        "problem": problem,
        "target": conn.handle,
        "number": conn.number,
        "drawing": conn.drawing.number,
        "location": (round(conn.location.x, 2), round(conn.location.y, 2))
    }


def number_matched(conn: Connector) -> bool:
    # return conn.number[:-2] == conn.drawing.number[1:]
    return int(conn.number[:-2]) == int(conn.drawing.number)


def is_excluded(conn: Connector) -> bool:
    # exclude legend drawings
    return int(conn.drawing.number) < 10


def check(p: PnID):
    connectors = p.main_connectors
    problems = []
    pretty = PrettyPrinter()
    for connector in connectors:
        if is_excluded(connector):
            continue
        if not connector.number:
            problems.append(problem_line(connector, "Missing number"))
        elif not (connector.is_to | connector.is_from):
            problems.append(problem_line(connector, "Missing route"))
        elif connector.is_entering != connector.is_from:
            problems.append(problem_line(connector, "Wrong direction"))
        elif connector.is_to & (not number_matched(connector)):
            problems.append(problem_line(connector, "Wrong number when exiting"))
        elif connector.is_off_drawing & connector.is_from & number_matched(connector):
            problems.append(problem_line(connector, "Wrong number when entering"))
        elif connector.is_off_boundary & bool(connector.link_drawing):
            problems.append(problem_line(connector, "P&ID No. not blank in off-boundary conn"))
        elif connector.is_off_drawing & (not bool(connector.link_drawing)):
            problems.append(problem_line(connector, "Missing P&ID No. in off-drawing conn"))
    print(f"{len(problems)} problems detected:")
    pretty.pprint(problems)
