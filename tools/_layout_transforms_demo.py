import win32com.client as w, time, os, json, shutil
DL = os.path.join(os.path.expanduser("~"), "Downloads", "decko_layouts")
SRC = os.path.join(DL, "_SOURCE.pptx")
CARRIER = os.path.abspath('PPT_AI_Editor.pptm')
NAVY = "#1F3864"; MID = "#2F5597"; LT = "#8FAADC"; PALE = "#D9E2F3"

def app_open():
    a = None
    for i in range(10):
        try:
            a = w.DispatchEx('PowerPoint.Application'); a.Visible = True; break
        except Exception:
            time.sleep(2)
    return a

# bullet texts the LLM would read out of the snapshot's paragraphs
B4 = [("1", "Accelerate enterprise ARR - focused land-and-expand in the top 200 named accounts"),
      ("2", "Ship the AI copilot to GA and drive attach across the installed base"),
      ("3", "Expand gross margin +200 bps via the multi-cloud cost-optimization program"),
      ("4", "Deepen the partner channel - ecosystem pipeline reaches one third of new bookings")]
B3 = [("1", "Cost discipline - take out 15% of run-rate opex; consolidate vendors & real estate"),
      ("2", "Revenue quality - shift the mix toward recurring software, away from low-margin services"),
      ("3", "Talent density - top-performer attrition below 8%; rebuild the senior engineering bench")]

def card(slide, ref, x, y, wd, h, fill, txt, fc="#FFFFFF", fs=14):
    return [{"type": "add_shape", "slide": slide, "kind": "rrect",
             "pos": {"left": x, "top": y, "width": wd, "height": h},
             "fill": fill, "text": txt, "font_color": fc, "font_size": fs,
             "h_align": "left", "v_align": "middle", "ref_name": ref}]

def tbox(slide, ref, x, y, wd, h, txt, fs=11, fc=NAVY, bold=False, ha="left"):
    return [{"type": "add_text_box", "slide": slide, "text": txt,
             "pos": {"left": x, "top": y, "width": wd, "height": h},
             "font_size": fs, "font_bold": bold, "font_color": fc, "h_align": ha, "ref_name": ref}]

def mv(slide, ref, x, y): return {"type": "move_shape", "slide": slide, "shape_id": ref, "left": x, "top": y}
def rs(slide, ref, wd, h): return {"type": "resize_shape", "slide": slide, "shape_id": ref, "width": wd, "height": h}
def mr(slide, ref, x, y, wd, h): return [mv(slide, ref, x, y), rs(slide, ref, wd, h)]

def T1():  # slide 1 bullets -> QUAD 2x2 cards
    fills = [NAVY, MID, LT, PALE]; fcs = ["#FFFFFF", "#FFFFFF", "#FFFFFF", NAVY]
    cells = [(40, 64, 432, 222), (488, 64, 432, 222), (40, 294, 432, 222), (488, 294, 432, 222)]
    acts = [{"type": "delete_shape", "slide": 1, "shape_id": "bulletbox"}]
    for i, ((n, t), (x, y, wd, h)) in enumerate(zip(B4, cells)):
        acts += card(1, "q%d" % (i + 1), x, y, wd, h, fills[i], "%s   %s" % (n, t), fcs[i], 15)
    acts.append({"type": "enable_text_shrink_for_overflow", "scope": "slide:1"})
    return ("bullets_to_quad", 1, acts)

def T2():  # slide 1 bullets -> 4 COLUMNS
    xs = [40, 263, 486, 709]; wd = 209; fills = [NAVY, MID, LT, PALE]; fcs = ["#FFFFFF", "#FFFFFF", "#FFFFFF", NAVY]
    acts = [{"type": "delete_shape", "slide": 1, "shape_id": "bulletbox"}]
    for i, (n, t) in enumerate(B4):
        acts += card(1, "c%d" % (i + 1), xs[i], 64, wd, 452, fills[i], "%s\n\n%s" % (n, t), fcs[i], 14)
    acts.append({"type": "enable_text_shrink_for_overflow", "scope": "slide:1"})
    return ("bullets_to_4col", 1, acts)

def T3():  # slide 2 (3 bullets) -> 1 BIG LEFT + 2 STACKED RIGHT
    acts = [{"type": "delete_shape", "slide": 2, "shape_id": "bulletbox3"}]
    acts += card(2, "big", 40, 64, 580, 452, NAVY, "%s\n\n%s" % (B3[0][0], B3[0][1]), "#FFFFFF", 20)
    acts += card(2, "r1", 636, 64, 284, 222, MID, "%s\n\n%s" % (B3[1][0], B3[1][1]), "#FFFFFF", 14)
    acts += card(2, "r2", 636, 294, 284, 222, LT, "%s\n\n%s" % (B3[2][0], B3[2][1]), "#FFFFFF", 14)
    acts.append({"type": "enable_text_shrink_for_overflow", "scope": "slide:2"})
    return ("bullets_to_1L_2R", 2, acts)

def T4():  # slide 3 mixed -> 67/33 L-R
    S = 3; acts = []
    acts += mr(S, "m_chart", 40, 64, 580, 200)
    acts += mr(S, "m_table", 40, 276, 580, 130)
    acts += mr(S, "m_box1", 636, 64, 284, 60)
    acts += mr(S, "m_box2", 636, 130, 284, 60)
    acts += mr(S, 8, 636, 206, 24, 24); acts += mr(S, "m_cap1", 666, 206, 254, 24)
    acts += mr(S, 11, 636, 238, 24, 24); acts += mr(S, "m_cap2", 666, 238, 254, 24)
    acts += mr(S, "m_foot", 40, 520, 880, 14)
    acts.append({"type": "add_line", "slide": S, "x1": 624, "y1": 60, "x2": 624, "y2": 512, "color": "#BFBFBF", "weight_pt": 1.0})
    acts.append({"type": "enable_text_shrink_for_overflow", "scope": "slide:3"})
    return ("mixed_67L_33R", S, acts)

def T5():  # slide 3 mixed -> 2 STACKED LEFT + 1 FULL RIGHT
    S = 3; acts = []
    acts += mr(S, "m_chart", 40, 64, 580, 222)
    acts += mr(S, "m_table", 40, 294, 280, 130)
    acts += mr(S, "m_box1", 328, 294, 292, 62)
    acts += mr(S, "m_box2", 328, 360, 292, 62)
    acts += tbox(S, "r_hdr", 636, 64, 284, 18, "HIGHLIGHTS", 11, NAVY, True)
    acts += mr(S, 8, 636, 90, 26, 26); acts += mr(S, "m_cap1", 668, 90, 252, 26)
    acts += mr(S, 11, 636, 124, 26, 26); acts += mr(S, "m_cap2", 668, 124, 252, 26)
    acts += mr(S, "m_foot", 40, 522, 880, 14)
    acts.append({"type": "add_line", "slide": S, "x1": 624, "y1": 60, "x2": 624, "y2": 516, "color": "#BFBFBF", "weight_pt": 1.0})
    acts.append({"type": "add_line", "slide": S, "x1": 40, "y1": 290, "x2": 620, "y2": 290, "color": "#BFBFBF", "weight_pt": 1.0})
    acts.append({"type": "enable_text_shrink_for_overflow", "scope": "slide:3"})
    return ("mixed_2Lstack_1Rfull", S, acts)

def T6():  # slide 3 mixed -> QUAD, table CONVERTED TO CHART in TR cell
    S = 3; acts = []
    acts += mr(S, "m_chart", 40, 64, 432, 222)
    acts.append({"type": "delete_shape", "slide": S, "shape_id": "m_table"})
    acts.append({"type": "add_chart", "slide": S, "chart_type": "columnclustered",
                 "pos": {"left": 488, "top": 64, "width": 432, "height": 222},
                 "categories": ["Gross %", "FCF ($M)", "NPS"],
                 "series": [{"name": "FY24", "values": [61.2, 88, 42], "color": NAVY},
                            {"name": "FY25", "values": [63.8, 121, 47], "color": LT}],
                 "title": "FY24 vs FY25 - Key KPIs", "show_legend": True, "clean_style": True, "ref_name": "m_kpichart"})
    acts.append({"type": "set_chart_legend", "slide": S, "shape_id": "m_kpichart", "props": {"position": "bottom", "font_size": 8}})
    acts.append({"type": "set_chart_series", "slide": S, "shape_id": "m_kpichart", "series_index": 1, "props": {"show_labels": True, "label_format": "#,##0.0", "label_size": 8}})
    acts.append({"type": "set_chart_series", "slide": S, "shape_id": "m_kpichart", "series_index": 2, "props": {"show_labels": True, "label_format": "#,##0.0", "label_size": 8}})
    acts.append({"type": "set_chart_gridlines", "slide": S, "shape_id": "m_kpichart", "props": {"major": False}})
    acts += mr(S, "m_box1", 40, 294, 432, 104)
    acts += mr(S, "m_box2", 40, 406, 432, 104)
    acts += tbox(S, "br_hdr", 488, 294, 432, 18, "OTHER HIGHLIGHTS", 11, NAVY, True)
    acts += mr(S, 8, 488, 322, 26, 26); acts += mr(S, "m_cap1", 520, 322, 400, 26)
    acts += mr(S, 11, 488, 358, 26, 26); acts += mr(S, "m_cap2", 520, 358, 400, 26)
    acts += mr(S, "m_foot", 40, 522, 880, 14)
    acts.append({"type": "enable_text_shrink_for_overflow", "scope": "slide:3"})
    return ("mixed_quad_table2chart", S, acts)

def T7():  # slide 3 mixed -> 50/50 TOP/BOTTOM
    S = 3; acts = []
    acts += mr(S, "m_chart", 40, 64, 560, 222)
    acts += mr(S, "m_table", 620, 64, 300, 222)
    acts += mr(S, "m_box1", 40, 300, 280, 104)
    acts += mr(S, "m_box2", 336, 300, 280, 104)
    acts += mr(S, 8, 632, 304, 26, 26); acts += mr(S, "m_cap1", 664, 304, 256, 26)
    acts += mr(S, 11, 632, 340, 26, 26); acts += mr(S, "m_cap2", 664, 340, 256, 26)
    acts += mr(S, "m_foot", 40, 522, 880, 14)
    acts.append({"type": "add_line", "slide": S, "x1": 40, "y1": 292, "x2": 920, "y2": 292, "color": "#BFBFBF", "weight_pt": 1.0})
    acts.append({"type": "enable_text_shrink_for_overflow", "scope": "slide:3"})
    return ("mixed_50T_50B", S, acts)

TRANSFORMS = [T1, T2, T3, T4, T5, T6, T7]

app = app_open()
carrier = app.Presentations.Open(CARRIER, WithWindow=True)
for fn in TRANSFORMS:
    name, slide, acts = fn()
    out = os.path.join(DL, name + ".pptx")
    shutil.copyfile(SRC, out)
    d = app.Presentations.Open(out, WithWindow=True)
    r = app.Run('PPT_AI_Editor!ExecuteFromString', json.dumps({"actions": acts}))
    d.Save()
    png = os.path.join(DL, name + ".png")
    d.Slides(slide).Export(png, 'PNG', 1280, 720)
    d.Close()
    open(os.path.join(DL, name + "_actions.json"), 'w', encoding='utf-8').write(json.dumps({"actions": acts}, indent=2))
    try:
        os.remove(out + ".action_log.jsonl")
    except Exception:
        pass
    print(name, "->", r)
carrier.Close(); app.Quit()
print("DONE")
