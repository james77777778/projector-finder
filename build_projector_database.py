import sys
import json

import pandas as pd


def read_projector(database):
    projector_dict = {}
    i = 1

    while True:
        # database.iloc[1][0]: name
        # database.iloc[1][1]: Brightness (lm)
        # database.iloc[1][2]: Native aspect ratio
        # database.iloc[1][3]: Native aspect rate
        # database.iloc[1][4]: Throw ratio (low)
        # database.iloc[1][5]: Throw ratio (high)
        # database.iloc[1][6]: Native resolution
        # database.iloc[1][7]: 距離上限 (m)

        # stop criterion
        if pd.isna(database.iloc[i][1]):
            print("Stop scan: there are {} items.".format(len(projector_dict)))
            break

        # parsing
        name = database.iloc[i][0]
        if name in projector_dict:
            raise ValueError("There is duplicate name: {}".format(name))
        projector_dict[name] = {}
        projector_dict[name]['brightness'] = float(database.iloc[i][1])
        # aspect_ratio = database.iloc[i][2].split(":")
        # projector_dict[name]['aspect_ratio'] = [float(x) for x in aspect_ratio]

        throw_ratio_low = float(database.iloc[i][4])
        throw_ratio_high = float(database.iloc[i][5])
        if pd.isna(database.iloc[i][5]):  # deal with no high throw_ratio
            throw_ratio_high = throw_ratio_low
        projector_dict[name]['throw_ratio'] = [throw_ratio_low, throw_ratio_high]
        projector_dict[name]['resolution'] = str(database.iloc[i][6])

        if pd.notna(database.iloc[i][7]):
            projector_dict[name]['distance_upper_bound'] = float(database.iloc[i][7])
        else:
            projector_dict[name]['distance_upper_bound'] = -1.0
        i += 1
    return projector_dict


if __name__ == "__main__":
    if len(sys.argv) != 2 or ".xlsx" not in sys.argv[1]:
        raise ValueError("You need to specify the input path of projector database excel file: sys.argv = {}".format(
            sys.argv
        ))
    database = pd.read_excel(sys.argv[1], engine='openpyxl')
    projector_dict = read_projector(database)

    for k, v in projector_dict.items():
        print(k, v)

    with open("projector_database.json", "w") as f:
        json.dump(projector_dict, f, indent=4)
