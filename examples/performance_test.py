from caddoc import CADDoc
import time

# A performance test for indexing blockref by traverse VS by selection
# Result as bellow:
# Current File: LNG.Plant.PnID_2021.0507B.dwg
# Indexing blockrefs...
# Indexing complete.
# 2736 collected by brute force, cost 42.989338874816895 s
# Indexing blockrefs...
# Indexing complete.
# 2736 collected by select, cost 9.911434888839722 s
if __name__ == "__main__":
    drawing = CADDoc(load_data=False)
    methods = {
        "brute force": False,
        "select": True
    }
    for method in methods:
        start_time = time.time()
        db = drawing.gen_blockref_dict(methods[method])
        time_cost = time.time() - start_time
        counter = sum((len(item) for item in db))
        print(f"{counter} collected by {method}, cost {time_cost} s")
