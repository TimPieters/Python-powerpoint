from pptx import Presentation
import os
from pptx.util import Inches, Pt
from pptx.enum.text import MSO_ANCHOR, MSO_AUTO_SIZE
from pptx.dml.color import RGBColor
from pptx.enum.dml import MSO_THEME_COLOR
from pptx.enum.text import PP_ALIGN
import pandas as pd
import numpy as np

df_xlsx = pd.read_excel("D:/Upthrust/Frank/PoC Cookbook Automation/Excel/Volkswagen Cookbook Excel.xlsx")






prs = Presentation()

title_slide_layout = prs.slide_layouts[0]
slide = prs.slides.add_slide(title_slide_layout)
title = slide.shapes.title
subtitle = slide.placeholders[1]


title.text = "VW Cookbook"
subtitle.text = "Growthmarketing"


for lab , cont in df_xlsx.iterrows() :

    a = cont
    header = a[0]
    started = str(a[1])
    status = a[2]
    channels = a[3]
    total_reach = str(a[4])
    total_clicks = str(a[5])
    media_spend = "€" + str(a[6])
    total_leads = str(a[7])
    cost_per_lead = "€" + str(a[8])
    ads_set_up = str(a[9])

    print(a)

    def main(i):

        normal_slide = prs.slide_layouts[5]
        slide_2 = prs.slides.add_slide(normal_slide)
        title = slide_2.shapes.title


        title.text = "Experiment 1: Search Campaign Per Model"



        for shape in slide_2.shapes:
            if not shape.has_text_frame:
                continue
            text_frame = shape.text_frame
            # do things with the text frame


        text_frame = shape.text_frame
        text_frame.clear()



        p = text_frame.paragraphs[0]
        run = p.add_run()
        run.text = header

        font = run.font
        font.name = 'Century Goth'
        font.size = Pt(25.3)
        font.bold = True
        font.italic = None  # cause value to be inherited from theme



        txBox = slide_2.shapes.add_textbox(Inches(0), Inches(6.5), Inches(2), Inches(1))
        tf = txBox.text_frame



        for shape in slide_2.shapes:
            if not shape.has_text_frame:
                continue
            text_frame = shape.text_frame
            # do things with the text frame


        text_frame = shape.text_frame

        p = text_frame.paragraphs[0]
        run = p.add_run()
        run.text = 'Started: '

        font = run.font
        font.name = 'Century Goth'
        font.size = Pt(14)
        font.bold = True
        font.italic = None  # cause value to be inherited from theme



        run = p.add_run()
        run.text = started

        font = run.font
        font.name = 'Century Goth'
        font.size = Pt(14)
        font.bold = False
        font.italic = None  # cause value to be inherited from theme

        p = tf.add_paragraph()



        run = p.add_run()
        run.text = 'Status: '

        font = run.font
        font.name = 'Century Goth'
        font.size = Pt(14)
        font.bold = True
        font.italic = None  # cause value to be inherited from theme


        run = p.add_run()
        run.text = status

        font = run.font
        font.name = 'Century Goth'
        font.size = Pt(14)
        font.bold = False
        font.italic = None  # cause value to be inherited from theme


        p = tf.add_paragraph()

        run = p.add_run()
        run.text = 'Channel(s): '

        font = run.font
        font.name = 'Century Goth'
        font.size = Pt(14)
        font.bold = True
        font.italic = None  # cause value to be inherited from theme


        run = p.add_run()
        run.text = channels

        font = run.font
        font.name = 'Century Goth'
        font.size = Pt(14)
        font.bold = False
        font.italic = None  # cause value to be inherited from theme


        txBox = slide_2.shapes.add_textbox(Inches(3), Inches(6.5), Inches(2), Inches(1))
        tf = txBox.text_frame

        for shape in slide_2.shapes:
            if not shape.has_text_frame:
                continue
            text_frame = shape.text_frame
            # do things with the text frame


        text_frame = shape.text_frame

        p = text_frame.paragraphs[0]
        run = p.add_run()
        run.text = 'Current Results: '

        font = run.font
        font.name = 'Century Goth'
        font.size = Pt(14)
        font.bold = True
        font.italic = None  # cause value to be inherited from theme

        txBox = slide_2.shapes.add_textbox(Inches(5), Inches(6.5), Inches(2), Inches(1))
        tf = txBox.text_frame


        for shape in slide_2.shapes:
            if not shape.has_text_frame:
                continue
            text_frame = shape.text_frame
            # do things with the text frame


        text_frame = shape.text_frame

        p = text_frame.paragraphs[0]
        run = p.add_run()
        run.text = 'Total Reach: '

        font = run.font
        font.name = 'Century Goth'
        font.size = Pt(14)
        font.bold = True
        font.italic = None  # cause value to be inherited from theme


        text_frame = shape.text_frame


        run = p.add_run()
        run.text = total_reach

        font = run.font
        font.name = 'Century Goth'
        font.size = Pt(14)
        font.bold = False
        font.italic = None  # cause value to be inherited from theme

        p = tf.add_paragraph()

        text_frame = shape.text_frame


        run = p.add_run()
        run.text = 'Total Clicks: '

        font = run.font
        font.name = 'Century Goth'
        font.size = Pt(14)
        font.bold = True
        font.italic = None  # cause value to be inherited from theme


        text_frame = shape.text_frame



        run = p.add_run()
        run.text = total_clicks

        font = run.font
        font.name = 'Century Goth'
        font.size = Pt(14)
        font.bold = False
        font.italic = None  # cause value to be inherited from theme

        p = tf.add_paragraph()

        text_frame = shape.text_frame


        run = p.add_run()
        run.text = 'Media Spend: '

        font = run.font
        font.name = 'Century Goth'
        font.size = Pt(14)
        font.bold = True
        font.italic = None  # cause value to be inherited from theme


        text_frame = shape.text_frame


        run = p.add_run()
        run.text = media_spend

        font = run.font
        font.name = 'Century Goth'
        font.size = Pt(14)
        font.bold = False
        font.italic = None  # cause value to be inherited from theme


        txBox = slide_2.shapes.add_textbox(Inches(7.5), Inches(6.5), Inches(2), Inches(1))
        tf = txBox.text_frame


        for shape in slide_2.shapes:
            if not shape.has_text_frame:
                continue
            text_frame = shape.text_frame
            # do things with the text frame


        text_frame = shape.text_frame

        p = text_frame.paragraphs[0]
        run = p.add_run()
        run.text = 'Total Leads: '

        font = run.font
        font.name = 'Century Goth'
        font.size = Pt(14)
        font.bold = True
        font.italic = None  # cause value to be inherited from theme


        text_frame = shape.text_frame


        run = p.add_run()
        run.text = total_leads

        font = run.font
        font.name = 'Century Goth'
        font.size = Pt(14)
        font.bold = False
        font.italic = None  # cause value to be inherited from theme

        p = tf.add_paragraph()

        text_frame = shape.text_frame


        run = p.add_run()
        run.text = 'Cost Per Lead: : '

        font = run.font
        font.name = 'Century Goth'
        font.size = Pt(14)
        font.bold = True
        font.italic = None  # cause value to be inherited from theme


        text_frame = shape.text_frame


        run = p.add_run()
        run.text = cost_per_lead

        font = run.font
        font.name = 'Century Goth'
        font.size = Pt(14)
        font.bold = False
        font.italic = None  # cause value to be inherited from theme


        p = tf.add_paragraph()

        text_frame = shape.text_frame


        run = p.add_run()
        run.text = 'Ads Set Up: : '

        font = run.font
        font.name = 'Century Goth'
        font.size = Pt(14)
        font.bold = True
        font.italic = None  # cause value to be inherited from theme


        text_frame = shape.text_frame

        run = p.add_run()
        run.text = ads_set_up

        font = run.font
        font.name = 'Century Goth'
        font.size = Pt(14)
        font.bold = False
        font.italic = None  # cause value to be inherited from theme
        prs.save('D:/Upthrust/Frank/PoC Cookbook Automation/Powerpoints/Test.pptx')
    main(25)



os.startfile("D:/Upthrust/Frank/PoC Cookbook Automation/Powerpoints/Test.pptx")









