from ..SapGui import SapGui


class Tree:
    def __init__(self, sap_gui: SapGui):
        self.sap_gui = sap_gui


    def _get_all_node_keys(self, field_id: str):
        return self.sap_gui.get_object(field_id).GetAllNodeKeys()
