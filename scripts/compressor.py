# Compress multiple layers to one for reference insertion, avoid breaking cleanness of the target drawing.

from caddoc import CADDoc


def compress(filepath: str = None):
    doc = CADDoc(filepath=filepath, load_data=False)
    entities = doc.select_all_drawing_objects()
    print(f"Starting with {len(entities)} drawing objects...")
    counter = 0
    for entity in entities:
        layer_name = entity.Layer
        layer = doc.doc.Layers.Item(layer_name)
        if layer_name != "0" and layer.LayerOn and (not layer.Freeze) and entity.Visible:
            if entity.Linetype == "BYLAYER":
                entity.Linetype = layer.Linetype
            # 192 means color by layer
            if entity.TrueColor.ColorMethod == 192:
                entity.TrueColor = layer.TrueColor
            entity.Layer = "0"
            # entity.TrueColor = color
            counter += 1
    print(f"Finish compress {counter} objects.")


if __name__ == "__main__":
    compress()
