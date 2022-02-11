from collections import defaultdict
from typing import Iterable

from components import Connector, MainConnector
from config import load_config
from pnid import PnID
from pprint import PrettyPrinter


def problem_line(connector: Connector, problem: str) -> dict:
    return {
        "problem": problem,
        # "target": conn.handle,
        "number": connector.tag,
        "drawing": connector.drawing.tag,
        "location": (round(connector.position.x, 2), round(connector.position.y, 2)),
    }


def number_matched(connector: Connector, config: dict) -> bool:
    """
    Check number convention
    """
    connector_prefix = connector.tag[:config["drawing"]["number_digits"]]
    return connector_prefix == get_dwg_number(connector, config)


def is_excluded(connector: Connector, config: dict) -> bool:
    # exclude legend drawings
    dwg_number = get_dwg_number(connector, config)
    return int(get_unit(dwg_number, config)) < config["drawing"]["start_unit"]


def get_dwg_number(connector: Connector, config: dict) -> str:
    digits = config["drawing"]["number_digits"]
    return connector.drawing.tag[-digits:]


def get_unit(number: str, config: dict) -> str:
    return number[:config["drawing"]["unit_digits"]]


def get_seq(number: str, config: dict) -> str:
    return number[config["drawing"]["unit_digits"]:]


def check(pnid: PnID, config: dict):
    pretty = PrettyPrinter()
    print("MainConnectors:")
    pretty.pprint(check_main(pnid, config))
    print("UtilityConnectors:")
    pretty.pprint(check_utility(pnid, config))


def check_main(pnid: PnID, config: dict) -> list:
    connectors = pnid.main_connectors
    problems = []
    for connector in connectors:
        try:
            if is_excluded(connector, config):
                continue
            if not connector.tag:
                problems.append(problem_line(connector, "Missing number"))
            elif not (connector.is_to | connector.is_from):
                problems.append(problem_line(connector, "Missing route"))
            elif connector.is_entering != connector.is_from:
                problems.append(problem_line(connector, "Wrong direction"))
            elif connector.is_to & (not number_matched(connector, config)):
                problems.append(problem_line(connector, "Wrong number when exiting"))
            elif connector.is_off_drawing & connector.is_from & number_matched(connector, config):
                problems.append(problem_line(connector, "Wrong number when entering"))
            elif connector.is_off_boundary & bool(connector.link_drawing):
                problems.append(problem_line(connector, "P&ID No. not blank in off-boundary connector"))
            elif connector.is_off_drawing & (not bool(connector.link_drawing)):
                problems.append(problem_line(connector, "Missing P&ID No. in off-drawing connector"))
        except KeyError as err:
            problems.append(problem_line(connector, str(err)))
    print(f"{len(problems)} problems detected:")
    return problems


def check_basic(connector: Connector) -> list:
    problems = []
    if not connector.tag:
        problems.append(problem_line(connector, "Missing tag number"))

    print(f"{len(problems)} problems detected:")
    return problems


def check_utility(pnid: PnID, config: dict) -> list:
    problems = []
    for connector in pnid.utility_connectors:
        if is_excluded(connector, config):
            continue
        if not connector.tag:
            problems.append(problem_line(connector, "Missing number"))

    return problems


def show_links(main_connectors: Iterable[MainConnector], config: dict):
    problems = []
    links_report = []
    links = defaultdict(list)
    # build db
    for connector in main_connectors:
        links[connector.tag].append(connector)
    # print(links)
    # make pair
    for tag in sorted(links):
        if len(links[tag]) < 3:
            start_connector = None
            end_connector = None
            for connector in links[tag]:
                if connector.is_from:
                    end_connector = connector
                elif connector.is_to:
                    start_connector = connector
            if start_connector and end_connector:
                start_info = f'[{get_dwg_number(start_connector, config)}]{end_connector.endpoint}'
                end_info = f'[{get_dwg_number(end_connector, config)}]{start_connector.endpoint}'
                links_report.append(f'{tag}: {start_info} -> {end_info}')
        else:
            problems.append(f'Duplicate number with link <{tag}>')

    return links_report


def report(pnid: PnID, config: dict):
    connectors = [connector for connector in pnid.main_connectors if not is_excluded(connector, config)]
    links = show_links(connectors, config)
    for link in links:
        print(link)


if __name__ == "__main__":
    pnid = PnID()
    config = load_config(r'..\config.ini')
    # check(pnid, config)
    report(pnid, config)
