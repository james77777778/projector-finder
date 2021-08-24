import pandas as pd


ambience = ["Only a small amount", "Light", "Bright"]
aspect_ratio = ["4:3", "16:9", "16:10"]
screen_size = ["<140inch", "140-199inch", "200-249inch", "250-300inch"]
distance = ["<2.0m", "2.0-3.5m", "3.5-5.5m", ">=5.5m"]
resolution = ["XGA", "WXGA", "1080P", "WUXGA", "4K"]

result = []
for amb in ambience:
    for aspect in aspect_ratio:
        for screen in screen_size:
            for dist in distance:
                for resol in resolution:
                    result.append(
                        (amb, aspect, screen, dist, resol)
                    )

output_result = pd.DataFrame(result)
output_result.to_excel("questions.xlsx", header=False)
