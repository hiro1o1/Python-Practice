"""Per-board configuration for the catalog image pipeline.

All coordinates are pixels in the matching ``catalog/source/board-N.webp``
image.  Each board describes:

* ``watermark`` – optional (alpha, rgb) of a semi-transparent overlay to undo.
* ``annotation_colors`` – which colored supplier annotations to strip
  (``red`` leader lines/dots, ``yellow`` label boxes).
* ``protect`` – points whose nearby red pixels are real board features
  (e.g. copper mounting rings) and must not be treated as annotations.
* ``erase`` – extra regions to inpaint:
  ``("rect", box)`` the whole box, ``("light", box)`` only light pixels
  (white silkscreen text), ``("dark", box)`` only dark pixels.
* ``arrows`` – supplier arrows as (tip, callout-circle-center) pairs.
* ``callouts`` – the new English labels.  ``side`` is one of
  top/bottom/left/right (label placed in the margin) or ``tag`` (label drawn
  on the board at ``tag``).  ``at`` is one anchor point or a list of them.
"""


def C(at, side, title, sub=None, tag=None):
    return {"at": at, "side": side, "title": title, "sub": sub, "tag": tag}


PUBLIC_TOUCH = "Public · USB touch pass-through"

BOARDS = [
    {
        "id": "rk3576-a",
        "source": "board-1.webp",
        "name": "RK3576 Mainboard A",
        "annotation_colors": ["red", "yellow"],
        "protect": [(768, 440), (1187, 440), (768, 857), (1187, 857)],
        "erase": [
            ("light", (700, 940, 985, 1000)),   # large model silkscreen
            ("light", (1110, 600, 1150, 740)),  # core-module silkscreen
            ("rect", (1262, 1054, 1352, 1106)),  # jack maker marking on RJ45
        ],
        "callouts": [
            C([(591, 234), (639, 235), (682, 235)], "top", "Antennas ANT1–ANT3",
              "Wi-Fi 2.4/5 GHz ×2 · Bluetooth"),
            C((779, 257), "top", "M.2 Wi-Fi Slot", "Optional"),
            C((998, 250), "top", "V-by-One"),
            C((1286, 249), "top", "OPS"),
            C((1524, 247), "top", "UART6 (TTL)", "Touchscreen UART"),
            C((1563, 252), "right", "USB 2.0", PUBLIC_TOUCH),
            C((1563, 321), "right", "USB 2.0", "Android · USB touch"),
            C((1563, 384), "right", "UART0 (TTL)", "Debug"),
            C((1541, 460), "right", "USB 2.0", PUBLIC_TOUCH),
            C((1550, 544), "right", "IR Remote"),
            C((1550, 582), "right", "Keypad"),
            C((1550, 620), "right", "LED Indicator"),
            C((1561, 685), "right", "Speaker", "Mid/high range"),
            C((1525, 722), "right", "Subwoofer", "Optional"),
            C((1560, 789), "right", "GPIO"),
            C((1561, 873), "right", "I²C"),
            C((1550, 983), "right", "Main Power Input"),
            C((1387, 1032), "right", "Audio Header"),
            C((453, 361), "left", "S/PDIF Out"),
            C((453, 555), "left", "UART8 (RS-232)"),
            C((521, 838), "left", "Front Panel", "Connector"),
            C((489, 1098), "bottom", "USB Type-B Out", "Touch pass-through"),
            C((617, 1098), "bottom", "HDMI In"),
            C((745, 1097), "bottom", "HDMI Out"),
            C((849, 1097), "bottom", "USB-C OTG"),
            C((948, 1097), "bottom", "USB 3.0", "Android"),
            C((1065, 1098), "bottom", "USB 2.0", "Android"),
            C((1182, 1098), "bottom", "USB 2.0", "Public"),
            C((1303, 1098), "bottom", "RJ45 Ethernet"),
            C((1406, 1098), "bottom", "Line Out"),
            C((1477, 1098), "bottom", "Line In"),
            C((570, 448), "tag", "RTC Battery", tag=(570, 405)),
            C((692, 446), "tag", "USB Expansion", tag=(712, 405)),
            C((868, 349), "tag", "Backlight", "Optional", tag=(868, 403)),
            C((657, 524), "tag", "External Wi-Fi Header", tag=(672, 594)),
            C((598, 720), "tag", "Power Header", "Optional", tag=(705, 722)),
            C((1053, 420), "tag", "USB Header", "Optional", tag=(1030, 382)),
            C((1102, 420), "tag", "UART Header", "Optional", tag=(1110, 330)),
            C((1504, 322), "tag", "USB Header", "Android · optional", tag=(1405, 325)),
            C((1503, 388), "tag", "USB Header", "Public · optional", tag=(1395, 398)),
        ],
    },
    {
        "id": "rk3576-b",
        "source": "board-2.webp",
        "name": "RK3576 Mainboard B",
        "annotation_colors": ["red", "yellow"],
        "protect": [(639, 478), (872, 478), (639, 711), (872, 711)],
        "erase": [
            ("light", (636, 824, 800, 866)),   # large model silkscreen
            ("light", (826, 570, 856, 656)),   # core-module silkscreen
        ],
        "callouts": [
            C([(505, 190), (526, 190), (546, 190), (566, 190), (587, 190), (609, 190)],
              "top", "Antennas ANT1–ANT7", "ANT1, ANT3, ANT7 optional"),
            C((726, 217), "top", "V-by-One"),
            C((928, 214), "top", "OPS"),
            C((1044, 212), "top", "USB 3.0", "Android"),
            C((404, 220), "left", "S/PDIF Out"),
            C((436, 256), "left", "M.2 Wi-Fi Slot", "Optional"),
            C((405, 311), "left", "RS-232"),
            C((414, 411), "left", "USB Type-B Out", "Touch pass-through"),
            C((399, 471), "left", "VGA Audio In"),
            C((409, 550), "left", "VGA In"),
            C((555, 578), "left", "Fast-Charge Power", "Connector"),
            C((402, 663), "left", "DisplayPort In"),
            C((396, 733), "left", "USB-C OTG", "65 W fast charge · DP In"),
            C((436, 849), "left", "Front Panel", "Connector"),
            C((1146, 253), "right", "USB 2.0", "Public"),
            C((1154, 312), "right", "UART6"),
            C((1154, 348), "right", "USB 2.0", "Android"),
            C((1144, 398), "right", "IR Remote"),
            C((1145, 426), "right", "Keypad"),
            C((1145, 452), "right", "LED Indicator"),
            C((1152, 503), "right", "I²C"),
            C((1135, 522), "right", "GPIO"),
            C((1142, 608), "right", "USB 2.0", "Public"),
            C((1142, 647), "right", "USB 2.0", "Android"),
            C((1142, 684), "right", "USB 2.0", "Android"),
            C((1153, 742), "right", "Speaker"),
            C((1153, 790), "right", "Subwoofer"),
            C((1156, 827), "right", "Subwoofer", "Optional"),
            C((1144, 896), "right", "Main Power Input"),
            C((1019, 950), "right", "Audio Header"),
            C((426, 988), "left", "HDMI In 1"),
            C((500, 990), "bottom", "HDMI In 2"),
            C((571, 989), "bottom", "USB Type-B", "Touch out"),
            C((647, 991), "bottom", "HDMI Out"),
            C((719, 991), "bottom", "USB 3.0", "Public"),
            C((758, 995), "bottom", "Recovery", "Button"),
            C((796, 989), "bottom", "USB 3.0", "Public"),
            C((873, 992), "bottom", "Ethernet 1", "RJ45"),
            C((955, 990), "bottom", "Ethernet 2", "RJ45 · optional"),
            C((1016, 991), "bottom", "Line Out"),
            C((1060, 986), "bottom", "Line In"),
            C((1116, 985), "right", "microSD", "Optional"),
            C((684, 266), "tag", "Backlight", "Optional", tag=(686, 298)),
            C([(611, 387), (631, 364)], "tag", "USB Expansion", "Optional", tag=(622, 416)),
            C((1117, 334), "tag", "USB / UART Header", "Optional", tag=(1003, 334)),
            C((991, 490), "tag", "RTC Battery", tag=(989, 459)),
            C((1147, 553), "tag", "RS-485", "Optional", tag=(1078, 522)),
            C((1123, 575), "tag", "UART9", "Optional", tag=(1052, 584)),
            C((1096, 644), "tag", "UART0", "Optional", tag=(1025, 644)),
        ],
    },
    {
        "id": "rk3588-a",
        "source": "board-3.webp",
        "name": "RK3588 Mainboard A",
        "watermark": (0.38, (235, 0, 0)),
        "erase": [
            ("rect", (376, 624, 510, 740)),     # maker logo + model silkscreen
            ("rect", (542, 634, 816, 694)),     # model sticker
            ("rect", (490, 686, 890, 748)),     # serial sticker + its outline
            ("white", (530, 926, 780, 1016)),   # backdrop watermark + callout rings
            ("white", (530, 898, 676, 926)),
        ],
        "arrows": [
            ((548, 882), (543, 978)), ((618, 888), (634, 978)),
            ((699, 890), (721, 978)), ((824, 906), (825, 978)),
            ((1210, 786), (1300, 787)), ((1205, 686), (1298, 686)),
            ((1213, 583), (1307, 585)), ((1220, 477), (1302, 479)),
            ((1197, 378), (1299, 380)), ((1223, 288), (1313, 300)),
            ((1223, 225), (1314, 225)), ((1193, 165), (1307, 156)),
            ((165, 129), (77, 129)),
        ],
        "callouts": [
            C((548, 870), "bottom", "Headphone Out"),
            C((620, 876), "bottom", "Mic In"),
            C((699, 880), "bottom", "S/PDIF Out", "Coaxial"),
            C((824, 896), "bottom", "RS-232"),
            C((1205, 786), "right", "USB Type-B", "Touch out"),
            C((1200, 686), "right", "HDMI In"),
            C((1205, 583), "right", "HDMI Out"),
            C((1210, 477), "right", "RJ45 Ethernet"),
            C((1190, 378), "right", "USB Type-C"),
            C((1215, 288), "right", "USB 2.0", "Android"),
            C((1215, 225), "right", "USB Host", "USB-A"),
            C((1185, 165), "right", "microSD Card"),
            C((175, 129), "left", "USB 2.0", "Public"),
        ],
    },
    {
        "id": "rk3588-b",
        "source": "board-4.webp",
        "name": "RK3588 Mainboard B",
        "watermark": (0.37, (240, 0, 0)),
        "erase": [
            ("rect", (1092, 846, 1230, 944)),   # maker logo + model silkscreen
            ("rect", (737, 966, 991, 1028)),    # model sticker
            ("rect", (518, 260, 862, 320)),     # blurred serial stickers
            ("rect", (1170, 798, 1227, 854)),   # QC sticker
        ],
        "arrows": [
            ((318, 1175), (276, 1272)), ((410, 1172), (386, 1277)),
            ((497, 1180), (496, 1285)), ((577, 1188), (604, 1285)),
            ((637, 1178), (714, 1290)), ((712, 1168), (825, 1290)),
            ((815, 1165), (934, 1293)), ((905, 1175), (1046, 1293)),
            ((1015, 1190), (1159, 1296)), ((1200, 1200), (1273, 1296)),
            ((1575, 1090), (1679, 1093)), ((1582, 1000), (1677, 1007)),
            ((1567, 900), (1675, 923)), ((1570, 802), (1672, 837)),
            ((1565, 707), (1667, 750)), ((1575, 612), (1669, 663)),
            ((1560, 502), (1669, 577)), ((1562, 400), (1672, 468)),
            ((1564, 306), (1672, 375)), ((1570, 248), (1669, 292)),
            ((1568, 194), (1667, 210)), ((1538, 136), (1667, 125)),
            ((272, 162), (164, 98)), ((268, 218), (162, 184)),
            ((272, 278), (162, 271)), ((272, 342), (162, 354)),
        ],
        "callouts": [
            C((318, 1170), "bottom", "RJ45 Ethernet", "Port 1"),
            C((410, 1168), "bottom", "RJ45 Ethernet", "Port 2"),
            C((497, 1175), "bottom", "Headphone Out"),
            C((577, 1182), "bottom", "Mic In"),
            C((637, 1172), "bottom", "S/PDIF Out", "Coaxial"),
            C((712, 1160), "bottom", "HDMI Out"),
            C((815, 1158), "bottom", "USB Type-B", "Touch out"),
            C((905, 1170), "bottom", "VGA Audio In"),
            C((1015, 1185), "bottom", "VGA In"),
            C((1200, 1195), "bottom", "RS-232"),
            C((1570, 1090), "right", "HDMI In 3"),
            C((1575, 1000), "right", "USB Type-B", "Touch for HDMI In 3"),
            C((1562, 900), "right", "HDMI In 2"),
            C((1565, 802), "right", "USB Type-B", "Touch for HDMI In 2"),
            C((1560, 707), "right", "HDMI In 1"),
            C((1570, 612), "right", "USB Type-B", "Touch for HDMI In 1"),
            C((1552, 502), "right", "DisplayPort In"),
            C((1555, 400), "right", "USB Type-C In"),
            C((1558, 306), "right", "USB 3.0", "Android"),
            C((1562, 248), "right", "USB 2.0", "Public"),
            C((1560, 194), "right", "USB 2.0", "Public"),
            C((1530, 136), "right", "microSD Card"),
            C((280, 162), "left", "USB 2.0", "Public"),
            C((276, 218), "left", "USB 2.0", "Public"),
            C((280, 278), "left", "USB 3.0", "Android"),
            C((280, 342), "left", "USB Touch", "Touch out"),
        ],
    },
]
