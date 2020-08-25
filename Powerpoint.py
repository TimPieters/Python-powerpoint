from pptx import Presentation
import os
from pptx.util import Inches, Pt
from pptx.enum.text import MSO_ANCHOR, MSO_AUTO_SIZE
from pptx.dml.color import RGBColor
from pptx.enum.dml import MSO_THEME_COLOR
from pptx.enum.text import PP_ALIGN
import pandas as pd
import numpy as np
from PIL import Image

df_xlsx = pd.read_excel("D:/Upthrust/Frank/PoC Cookbook Automation/Excel/Volkswagen Cookbook Excel.xlsx")

df_xlsx['Started'] = df_xlsx['Started'].dt.strftime('%m/%d/%Y')

prs = Presentation("D:/Upthrust/Frank/PoC Cookbook Automation/Powerpoints/Ok.pptx")
#16:9 lege powerpoint

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
    picture_1 = str(a[11])
    picture_2 = str(a[12])
    picture_3 = str(a[13])
    picture_4 = str(a[14])
    gray_border = "D:\\Upthrust\\Frank\\PoC Cookbook Automation\\Excel\\Pictures\\gray_border.png"
    arrow = "D:\\Upthrust\\Frank\\PoC Cookbook Automation\\Excel\\Pictures\\arrow.png"


    def main(i):

        normal_slide = prs.slide_layouts[5]
        slide_2 = prs.slides.add_slide(normal_slide)
        title = slide_2.shapes.title


        title.text = "Experiment 1: Search Campaign Per Model"

        img_path = gray_border

        top = Inches(6.5)
        left = Inches(0)
        height = Inches(1.1)
        pic = slide_2.shapes.add_picture(img_path, left, top, height=height)
        #een grijze border, puur estetisch.

        img_path = arrow

        top = Inches(3)
        left = Inches(6.5)
        height = Inches(2)
        pic = slide_2.shapes.add_picture(img_path, left, top, height=height)
        #een grijze border, puur estetisch.

        if picture_1 != "nan":

            im = Image.open(picture_1)
            width, height = im.size
            threshold = width - height

            img_path = picture_1

            if threshold < 100:
                top = Inches(3)
                left = Inches(0.5)
                height = Inches(2)

            else:
                top = Inches(2)
                left = Inches(0.5)
                height = Inches(2)


            pic = slide_2.shapes.add_picture(img_path, left, top, height=height)



        if picture_2 != "nan":

            im = Image.open(picture_2)
            width, height = im.size
            threshold = width - height

            img_path = picture_2

            if threshold < 100:
                top = Inches(3)
                left = Inches(3)
                height = Inches(2)
            else:
                top = Inches(4)
                left = Inches(0.5)
                height = Inches(2)

            pic = slide_2.shapes.add_picture(img_path, left, top, height=height)



        if picture_3 != "nan":

            im = Image.open(picture_3)
            width, height = im.size
            threshold = width - height

            img_path = picture_3

            if threshold < 100:
                top = Inches(2)
                left = Inches(7.5)
                height = Inches(4)

            else:
                top = Inches(2)
                left = Inches(7.5)
                height = Inches(4)


            pic = slide_2.shapes.add_picture(img_path, left, top, height=height)



        if picture_4 != "nan":

            im = Image.open(picture_4)
            width, height = im.size
            threshold = width - height

            img_path = picture_4

            if threshold < 100:
                top = Inches(3.5)
                left = Inches(10)
                height = Inches(1.2)
            else:
                top = Inches(3.5)
                left = Inches(10)
                height = Inches(1.2)

            pic = slide_2.shapes.add_picture(img_path, left, top, height=height)

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



        txBox = slide_2.shapes.add_textbox(Inches(0.5), Inches(6.5), Inches(2), Inches(1))
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


        txBox = slide_2.shapes.add_textbox(Inches(4), Inches(6.5), Inches(2), Inches(1))
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

        txBox = slide_2.shapes.add_textbox(Inches(7), Inches(6.5), Inches(2), Inches(1))
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


        txBox = slide_2.shapes.add_textbox(Inches(10), Inches(6.5), Inches(2), Inches(1))
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









