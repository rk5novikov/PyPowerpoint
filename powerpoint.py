import os
from win32com.client import Dispatch

from linalgebra import length, unit, summ

from xl_constants import *
from xl_constants_mso import *


rgb_colors = {
        'black'  : [  0,  0,  0], 
        'white'  : [255,255,255], 
        'red'    : [255,  0,  0], 
        'lime'   : [  0,255,  0], 
        'blue'   : [  0,  0,255], 
        'yellow' : [255,255,  0], 
        'cyan'   : [0  ,255,255], 
        'magenta': [255,0  ,255], 
        'silver' : [192,192,192], 
        'gray'   : [128,128,128], 
        'maroon' : [128,  0,  0], 
        'olive'  : [128,128,  0], 
        'green'  : [  0,128,  0], 
        'purple' : [128,  0,128], 
        'teal'   : [  0,128,128], 
        'navy'   : [  0,  0,128]
        }


BLANK_SLIDE_LAYOUT_ID = 2
TITLE_SLIDE_LAYOUT_ID = 1

text_orientation_constants = { 'horizontal': 1, 
                               'vertical': 5 }

default_text_options = { 'font_size': 12, 
                              'bold': False, 
                            'italic': False, 
                         'underline': False, 
                        'text_color': 'black', 
                              'fill': False, 
                        'fill_color': 'yellow' , 
                            'border': False, 
                      'border_color': 'black', 
                     'border_weight': 1.5 }

default_rectangle_options = {     'fill': False, 
                            'fill_color': 'yellow' , 
                                'border': True, 
                          'border_color': 'black', 
                         'border_weight': 1.5 }

default_line_options = { 'line_color': 'black', 
                        'line_weight': 1.5, 
                            'arrow_1': False, 
                            'arrow_2': False, 
                       'arrow_1_size': 1.0, 
                       'arrow_2_size': 1.0}


app = None


def rgb(rgb_int_list):
    return 65536*rgb_int_list[2] + 256*rgb_int_list[1] + 1*rgb_int_list[0]


def open_powerpoint(visible=True):
    global app
    app = Dispatch("PowerPoint.Application")
    app.Visible = visible
    return app


def open_document(doc_path):
    if doc_path:
        doc_path_norm = os.path.normpath(os.path.abspath(doc_path))
        cur_docs_fullpaths = [os.path.normpath(app.Presentations[i].FullName) for i in range(app.Presentations.Count)]
        if doc_path_norm in cur_docs_fullpaths:
            for i in range(app.Presentations.Count):
                if os.path.normpath(app.Presentations[i].FullName) == doc_path_norm:
                    app.Presentations[i].Windows(1).Activate
                    break
        else:
            app.Presentations.Open(doc_path)
    return app.ActivePresentation


def create_document(doc_path):
    new_document = None
    if doc_path:
        pass
    else:
        pass
    return new_document


def save():
    cur_ppt = app.ActivePresentation
    if not cur_ppt.Saved:
        cur_ppt.Save()
    return cur_ppt.FullName


def save_as(doc_path):
    cur_ppt = app.ActivePresentation
    if doc_path:
        cur_ppt = app.ActivePresentation
        try:
            if os.path.normpath(os.path.abspath(doc_path)) != os.path.normpath(os.path.abspath(app.ActivePresentation.FullName)):
                cur_ppt.SaveAs(os.path.abspath(doc_path))
            else:
                cur_ppt.Save()
        except:
            print("can't save under name '{0}'.pptx. Please verify that this document isnot already open now.".format(doc_path))
    return os.path.abspath(doc_path)


def quit_powerpoint():
    app.Quit()


def add_slide(pos='end', title='', layout_id=BLANK_SLIDE_LAYOUT_ID):
    slides = app.ActivePresentation.Slides
    pos_id = slides.Count + 1 if pos == 'end' else int(pos)
    slide = slides.Add(pos_id, Layout=layout_id)
    # delete empty text field
    text_boxes = [shape for shape in slide.Shapes if shape.Name.startswith('Text Placeholder')]
    if text_boxes:
        text_box = text_boxes[0]
        text_box.Delete()
    if title:
        slide_title_shapes = [ shape for shape in slide.Shapes if 'Title' in shape.Name ]
        if slide_title_shapes:
            slide_title_shape = slide_title_shapes[0]
            slide_title_shape.TextFrame.TextRange.Text = title
    slide.Select()
    return slide


def add_slide_title(pos='end', title='', layout_id=TITLE_SLIDE_LAYOUT_ID):
    slides = app.ActivePresentation.Slides
    slide_num = slides.Count
    slide = app.ActivePresentation.Slides[slide_num-1]
    pos_id = slides.Count + 1 if pos == 'end' else int(pos)
    if title:
        slide_title_shapes = [ shape for shape in slide.Shapes if 'Title' in shape.Name ]
        if slide_title_shapes:
            slide_title_shape = slide_title_shapes[0]
            slide_title_shape.TextFrame.TextRange.Text = title
    slide = slides.Add(pos_id, Layout=layout_id)
    slide.Select()
    return slide


def activate_slide(slide_num='end'):
    slides = app.ActivePresentation.Slides
    slide_num = slides.Count if slide_num=='end' else int(slide_num)
    slide = app.ActivePresentation.Slides[slide_num-1]
    slide.Select()
    return slide


def get_active_slide():
    return app.ActivePresentation.Slides(app.ActiveWindow.Selection.SlideRange.SlideIndex)


def get_slide_dimensions():
    page_setup = app.ActivePresentation.PageSetup
    return (page_setup.SlideWidth, page_setup.SlideHeight)


def apply_shape_position(shape, left, top, width, height):
    w_slide, h_slide = get_slide_dimensions()
    if left:
        shape.Left = left * w_slide
    if top:
        shape.Top = top * h_slide
    # 
    if width or height:
        w_pic, h_pic = shape.Width, shape.Height
        aspect = w_pic / h_pic
        if width and not height:
            shape.Width = width * w_slide
            shape.Height = shape.Width / aspect
        elif height and not width:
            shape.Width = shape.Width * aspect
            shape.Height = height * h_slide
        else:
            shape.Width = width * w_slide
            shape.Height = height * h_slide
    return shape


def apply_fill_border_options(shape, options_dict):
    if options_dict['fill']:
        shape.Fill.ForeColor.RGB = rgb(rgb_colors[options_dict['fill_color']])
    else:
        shape.Fill.Visible=False
    if options_dict['border']:
        shape.Line.DashStyle = 1
        shape.Line.ForeColor.RGB = rgb(rgb_colors[options_dict['border_color']])
        shape.Line.Weight = options_dict['border_weight']
    else:
        shape.Line.Visible=False


def insert_textbox(text, slide_id=None, orientation='horizontal', text_options={'font_size': 12}, left=None, top=None, width=None, height=None):
    _text_options = {k: v for k, v in default_text_options.items()}
    _text_options.update(text_options)
    if slide_id:
        activate_slide(slide_id)
    slide = get_active_slide()
    shapes = slide.Shapes
    if orientation in text_orientation_constants:
        mso_orentation = text_orientation_constants[orientation] 
    else:
        mso_orentation = text_orientation_constants['horizontal']
    textbox = shapes.AddTextbox(mso_orentation, Left=0, Top=0, Width=200, Height=20)
    textbox.TextFrame.AutoSize = 1 # ppAutoSizeShapeToFitText
    textbox.TextFrame.TextRange.Text = text
    # 
    font = textbox.TextFrame.TextRange.Font
    font.Bold = _text_options['bold']
    font.Italic = _text_options['italic']
    font.Size = _text_options['font_size']
    font.Underline = _text_options['underline']
    font.Color = rgb(rgb_colors[_text_options['text_color']])
    # 
    apply_fill_border_options(textbox, _text_options)
    apply_shape_position(textbox, left, top, width, height)
    return textbox


def insert_rectangle(slide_id=None, left=None, top=None, width=None, height=None, rectangle_options=default_rectangle_options):
    _rectangle_options = {k: v for k, v in default_rectangle_options.items()}
    _rectangle_options.update(rectangle_options)
    if slide_id:
        activate_slide(slide_id)
    slide = get_active_slide()
    shapes = slide.Shapes
    rectangle = shapes.AddShape(Type=1, Left=1, Top=1, Width=5, Height=5) # Type=msoShapeRectangle
    apply_fill_border_options(rectangle, _rectangle_options)
    apply_shape_position(rectangle, left, top, width, height)
    return rectangle



def insert_line(slide_id=None, begin_x=None, begin_y=None, vector=[0, 1], line_length=0.1, line_options=default_line_options):
    _line_options = {k: v for k, v in default_line_options.items()}
    _line_options.update(line_options)
    if slide_id:
        activate_slide(slide_id)
    slide = get_active_slide()
    shapes = slide.Shapes
    # 
    w_slide, h_slide = get_slide_dimensions()
    unit_vector = unit(vector)
    dim_coef = w_slide * line_length
    line = shapes.AddLine(BeginX=0, BeginY=0, EndX=unit_vector[0]*dim_coef, EndY=unit_vector[1]*dim_coef)
    line.Line.DashStyle = 1
    line.Line.ForeColor.RGB = rgb(rgb_colors[_line_options['line_color']])
    line.Line.Weight = _line_options['line_weight']
    if _line_options['arrow_1']:
        line.Line.BeginArrowheadStyle = 3
        line.Line.BeginArrowheadLength = _line_options['arrow_1_size']
    if _line_options['arrow_2']:
        line.Line.EndArrowheadStyle = 3
        line.Line.EndArrowheadLength = _line_options['arrow_2_size']
    apply_shape_position(line, begin_x, begin_y, width=None, height=None)
    return line


def insert_picture(file_path, slide_id=None, left=None, top=None, width=None, height=None, send_to_back=False):
    if slide_id:
        activate_slide(slide_id)
    slide = get_active_slide()
    shapes = slide.Shapes
    picture = shapes.AddPicture(FileName=os.path.abspath(file_path), LinkToFile=0, SaveWithDocument=1, Left=0, Top=0, Width=500, Height=500)
    picture.Scaleheight(1.0, 1)
    picture.Scalewidth(1.0, 1)
    picture = apply_shape_position(picture, left, top, width, height)
    if send_to_back:
        picture.ZOrder(1) # msoSendToBack
    return picture

