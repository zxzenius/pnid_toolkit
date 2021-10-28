from connector import Connector
from pnid import PnID
from pprint import PrettyPrinter


def problem_line(conn: Connector, problem: str) -> dict:
    return {
        "problem": problem,
        "target": conn.handle,
        "number": conn.tag,
        "drawing": get_dwg_num(conn),
        "location": (round(conn.position.x, 2), round(conn.position.y, 2))
    }


def number_matched(conn: Connector) -> bool:
    # return conn.number[:-2] == conn.drawing.number[1:]
    return int(conn.tag[:-2]) == int(get_dwg_num(conn))


def is_excluded(conn: Connector) -> bool:
    # exclude legend drawings
    return int(get_dwg_num(conn)) < 10


def get_dwg_num(conn: Connector) -> str:
    return conn.drawing.id[-4:]


def check(p: PnID):
    pretty = PrettyPrinter()
    print("MainConnectors:")
    pretty.pprint(check_main(p))
    print("UtilityConnectors:")
    pretty.pprint(check_utility(p))


def check_main(p: PnID) -> list:
    connectors = p.main_connectors
    problems = []
    for connector in connectors:
        if is_excluded(connector):
            continue
        if not connector.tag:
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
    return problems


def check_basic(connector: Connector) -> list:
    problems = []
    if not connector.tag:
        problems.append(problem_line(connector, "Missing tag number"))

    print(f"{len(problems)} problems detected:")
    return problems


def check_utility(p: PnID) -> list:
    problems = []
    for connector in p.utility_connectors:
        if is_excluded(connector):
            continue
        if not connector.tag:
            problems.append(problem_line(connector, "Missing number"))

    return problems


if __name__ == "__main__":
    p = PnID()
    check(p)
