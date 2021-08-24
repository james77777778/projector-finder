import re
import sys
import json

import pandas as pd


INPUT_PROJECTOR_DATABASE = "projector_database.json"
SHEET_NAME = "New matches"
OUTPUT_FILE = "result.xlsx"


def is_overlap(x1, x2, y1, y2):
    if y2 >= x1 and x2 >= y1:
        return True
    else:
        return False


def compute_screen_area(distance, aspect_ratio, throw_ratio_low):
    width = distance / throw_ratio_low
    height = width / (aspect_ratio[0] / aspect_ratio[1])

    area = (width ** 2 + height ** 2) ** (1 / 2)
    area_inch = area * 39.3701

    return area_inch


def compute_resolution(resolution):
    if resolution.upper() == "XGA":  # 1024x768
        return 0
    elif resolution.upper() == "WXGA":  # 1280x768
        return 1
    elif resolution.upper() == "1080P":  # 1920x1080
        return 2
    elif resolution.upper() == "WUXGA":  # 1920x1200
        return 3
    elif resolution.upper() == "4K":  # 3840×2160
        return 4
    else:
        raise NotImplementedError(
            "{} is unknown resolution.".format(resolution)
        )


def compute_ambience(ambience, screen_size):
    if ambience.lower() == "only a small amount":
        return [0.0, 4000.0]
    elif ambience.lower() == "light":
        if screen_size[1] <= 140:
            return [0.0, 4000.0]
        elif screen_size[0] == 140 and screen_size[1] == 199:
            return [4000.0, 5000.0]
        elif screen_size[0] == 200 and screen_size[1] == 249:
            return [5000.0, 6000.0]
        elif screen_size[0] == 250 and screen_size[1] == 300:
            return [6000.0, 999999.9]
        else:
            raise NotImplementedError(
                "{} is unknown screen size.".format(screen_size)
            )
    elif ambience.lower() == "bright":
        if screen_size[1] < 200:
            return [5000.0, 6000.0]
        else:
            return [6000.0, 999999.9]
    else:
        raise NotImplementedError(
            "{} is unknown ambience level.".format(ambience)
        )


def parse_question(question):
    # "題目", "Ambience", "Aspect Ratio", "Screen Size", "Distance"
    # "Projector Resolution", "此題有完全匹配的答案", "造成沒有完全匹配的選項"
    # "推薦產品1", "2", "3", "4"
    res = {}
    # ambience (Only a small amount, Light, Bright)
    ambience = question["Ambience"].strip()

    # Aspect Ratio (two float numbers)
    ratio = re.findall(r'\d+', question["Aspect Ratio"])
    aspect_ratio = [float(x) for x in ratio]

    # Screen Size (two int numbers)
    screen_size = re.findall(r'\d+', question["Screen Size"])
    screen_size = [int(x) for x in screen_size]
    if "<" in question["Screen Size"]:
        screen_size.insert(0, 0)

    # Distance (two float numbers)
    distance = re.findall(r"[-+]?\d*\.\d+|\d+", question["Distance"])
    distance = [float(x) for x in distance]
    if "<" in question["Distance"]:
        distance.insert(0, 0.0)
    if ">" in question["Distance"]:
        distance.append(100.0)

    # Resolution (1080P, 4K, WUXGA)
    resolution = question["Projector Resolution"].strip()

    res["Ambience"] = ambience
    res["Aspect Ratio"] = aspect_ratio
    res["Screen Size"] = screen_size
    res["Distance"] = distance
    res["Projector Resolution"] = resolution

    return res


def finder(question, projector_database):
    res = {
        "best match": [],
        "good match": []
    }

    target_screen_size = question["Screen Size"]
    target_resolution = question["Projector Resolution"]
    target_ambience = question["Ambience"]

    aspect_r = question["Aspect Ratio"]
    min_dis, max_dis = question["Distance"]

    # reason
    reason = []

    for projector_name, projector_info in projector_database.items():
        target_fails = []
        throw_ratio = projector_info["throw_ratio"]
        projector_resolution = projector_info["resolution"]
        projector_ambience = projector_info["brightness"]
        projector_dis_upper = projector_info["distance_upper_bound"]
        # check distance -> screen size
        is_distance_match = False
        # for min_distance -> screen size (low_throw_ratio, high_throw_ratio)
        min_screen_size = [
            # small value
            compute_screen_area(min_dis, aspect_r, throw_ratio[0]),
            # big value
            compute_screen_area(min_dis, aspect_r, throw_ratio[1])
        ]
        if is_overlap(min_screen_size[1], min_screen_size[0], target_screen_size[0], target_screen_size[1]):
            is_distance_match = True
        # for max_distance -> screen size (low_throw_ratio, high_throw_ratio)
        if max_dis != 0:
            max_screen_size = [
                compute_screen_area(max_dis, aspect_r, throw_ratio[0]),
                compute_screen_area(max_dis, aspect_r, throw_ratio[1])
            ]
            if is_overlap(max_screen_size[1], max_screen_size[0], target_screen_size[0], target_screen_size[1]):
                is_distance_match = True
        # check distance upper bound
        if projector_dis_upper > 0 and projector_dis_upper < min_dis:
            is_distance_match = False
        if not is_distance_match:
            target_fails.append("Distance")
        # check resolution
        if compute_resolution(projector_resolution) != compute_resolution(target_resolution):
            target_fails.append("Resolution")
        # check ambience
        target_ambience_ = compute_ambience(target_ambience, target_screen_size)
        if projector_ambience < target_ambience_[0] or projector_ambience > target_ambience_[1]:
            target_fails.append("Ambience")
        # conclusion
        if len(target_fails) == 0:  # best match
            res["best match"].append(projector_name)
        elif len(target_fails) == 1:  # good match
            res["good match"].append(projector_name)
            reason.append(target_fails[0])
            # print(target_fails)
    return res, reason


if __name__ == "__main__":
    if len(sys.argv) != 2 or ".xlsx" not in sys.argv[1]:
        raise ValueError("You need to specify the input path of projector matching excel file: sys.argv = {}".format(
            sys.argv
        ))
    projector_finder_database = pd.read_excel(
        sys.argv[1],
        sheet_name=SHEET_NAME,
        header=0,
        skiprows=lambda x: x in range(1, 7),
        engine='openpyxl'
    )

    with open(INPUT_PROJECTOR_DATABASE, "r") as f:
        projector_database = json.load(f)

    output_result = []
    for index, row in projector_finder_database.iterrows():
        if pd.notna(row["題目"]):  # only deal with questions
            # parse question
            question = parse_question(row)
            # find
            res, reason = finder(question, projector_database)
            output_result.append((res['best match'], res["good match"], reason))
            print(row["題目"], res)

    output_df = pd.DataFrame.from_records(output_result)
    output_df[0] = output_df[0].astype(str).str.replace('\\[|\\]|\'', '')
    output_df[1] = output_df[1].astype(str).str.replace('\\[|\\]|\'', '')
    output_df[2] = output_df[2].astype(str).str.replace('\\[|\\]|\'', '')
    output_df.to_excel(OUTPUT_FILE, header=False, index=False)
