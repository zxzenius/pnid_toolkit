# text style: eng <-> chs
from pnid import PnID


def switch_to_chs(p: PnID):
    """
    Switch language style from english to chinese.
    Including text style and some translation
    :param p:
    :return:
    """
    print("-> CHS")
    print("Start with connectors")
    counter = 0
    for connector in p.main_connectors + p.utility_connectors:
        # Exclude symbol legend
        if int(connector.drawing.tag) > 10:
            set_to_chs(connector.service_attr)
            if hasattr(connector, "route_attr"):
                set_to_chs(connector.route_attr)
                # translation
                connector.route = connector.route.replace("TO ", "至 ")
                connector.route = connector.route.replace("FROM ", "自 ")
            counter += 1

    print(f"{counter} connectors switched.")


def set_to_chs(attr_ref):
    attr_ref.StyleName = "ConnectorText_CHS"
    attr_ref.Height = 4
    attr_ref.ScaleFactor = 1


if __name__ == "__main__":
    p = PnID()
    switch_to_chs(p)
