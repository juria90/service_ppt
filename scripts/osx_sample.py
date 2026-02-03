"""Sample script for macOS PowerPoint automation.

This module contains example code demonstrating how to use the powerpoint_osx
module to automate PowerPoint operations on macOS using AppleScript.
"""

import os
from pathlib import Path

import service_ppt.powerpoint_osx as PowerPoint


def ppt_applescript():
    is_already_running = PowerPoint.App.is_running()

    powerpoint = PowerPoint.App()

    # prs = powerpoint.new_presentation()
    # prs.close()

    home = str(Path.home())
    templ_dir = os.path.join(home, "Church", "OnlineService")
    prs = powerpoint.open_presentation(os.path.join(templ_dir, "2020-Sunday-Template.pptx"))

    insert_location = 2
    prs.insert_blank_slide(insert_location)
    prs.delete_slide(insert_location)

    # duplicate_slides uses 'System Events' that requires Accessibility permission.
    # prs.duplicate_slides(9, 3, 5)

    print("slide count = %d" % prs.slide_count())

    slide_index = 0
    slide_id = prs.slide_index_to_ID(slide_index)
    print("slide index(%d) => slide ID(%d)" % (slide_index, slide_id))

    slide_index = prs.slide_ID_to_index(slide_id)
    print("slide ID(%d) => slide index(%d)" % (slide_id, slide_index))

    prs.replace_one_slide_texts(9, {"MK_service_title": "주일 예배", "MK_sermon_title": "유월절", "MK_date": "2020년 6월 30일"})

    # prs.replace_all_slides_texts({'MK_service_title': '주일 예배', 'MK_sermon_title': '유월절', 'MK_date': '2020년 6월 30일'})

    prs.insert_file_slides(10, os.path.join(templ_dir, "ppt", "NHymn304h.pptx"))

    prs.saveas(os.path.join(templ_dir, "2020-Sunday-temp.pptx"))

    output_dir = os.path.join(templ_dir, "output")
    prs.saveas_format(output_dir, "png")

    prs.close()

    powerpoint.quit_if_empty()


if __name__ == "__main__":
    ppt_applescript()
