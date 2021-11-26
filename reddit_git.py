from docx.enum.style import WD_STYLE_TYPE
from docx import Document
from docx.shared import Pt, RGBColor, Inches
from docx.oxml.ns import qn
import praw
import os
import sys
import dotenv
from dotenv import load_dotenv
load_dotenv()

# require('dotenv').config();

# console.log(process.env);


reddit = praw.Reddit(client_id=os.environ.get("CLIENT_ID"),
                     client_secret=os.environ.get("CLIENT_SECRET"),
                     username=os.environ.get("USER_NAME"),
                     password=os.environ.get("PASSWORD"),
                     user_agent='my user agent')


document = Document()
sections = document.sections
for section in sections:
    section.top_margin = Inches(0.5)
    section.bottom_margin = Inches(1)
    section.left_margin = Inches(1)
    section.right_margin = Inches(1)

section = document.sections[0]

# sectPr = section._sectPr
# cols = sectPr.xpath('./w:cols')[0]
# cols.set(qn('w:num'), '1')


def writedocx(content, font_name="Times New Roman", font_size=12, font_bold=False, font_italic=False, font_underline=False, color=RGBColor(0, 0, 0),
              before_spacing=5, after_spacing=5, line_spacing=1.5, keep_together=True, keep_with_next=False, page_break_before=False, style=""):
    paragraph = document.add_paragraph(str(content))
    # paragraph.style = document.styles.add_style(
    #     style, WD_STYLE_TYPE.PARAGRAPH)
    font = paragraph.style.font
    font.name = font_name
    font.size = Pt(font_size)
    font.bold = font_bold
    font.italic = font_italic
    font.underline = font_underline
    font.color.rgb = RGBColor(0, 0, 0)

    paragraph_format = paragraph.paragraph_format
    paragraph_format.space_before = Pt(12)
    paragraph_format.space_after = Pt(8)

    paragraph.line_spacing = line_spacing
    paragraph_format.keep_together = keep_together
    paragraph_format.keep_with_next = keep_with_next
    paragraph_format.page_break_before = page_break_before


number = 0
# with open('test.txt', 'w', encoding='utf-8') as f:
#     for item in reddit.user.me().saved(limit=None):
#         if isinstance(item, praw.models.Submission):
#             f.write('\n' + 'This is a post' + '\n')
#             # f.write(item.id + '\n')
#             f.write(item.title + '\n')
#             if item.is_self:
#                 f.write(item.selftext + '\n')
#             else:  # link post
#                 f.write(item.url)
#             f.write('-----------------------------------------------------')
#         else:  # comment
#             f.write('\n' + 'This is a comment' '\n')
#             f.write(item.id + '\n')
#             f.write(item.body + '\n')
#             print(item.subreddit)
#             if item.is_self:

#             f.write('-----------------------------------------------------')
#         number = number + 1
for item in reddit.user.me().saved(limit=None):
    writedocx(content='Subreddit: ' + str(item.subreddit),
              font_size=18, font_bold=True)
    if isinstance(item, praw.models.Submission):    # This means it's a post
        writedocx(content='\n' + 'This is a post' + '\n')
        writedocx(content=item.title, font_size=16, font_bold=True)
        if item.is_self:
            writedocx(content="BODY:    " + item.selftext + '\n')
        writedocx(content=item.url)
    else:                                           # This mean it's a comment
        writedocx(content="COMMENT BODY:    " + item.body + '\n')
    number = number + 1

document.save('D:\\stuff\\' +"reddit_word_experiment.docx")
print(number)
