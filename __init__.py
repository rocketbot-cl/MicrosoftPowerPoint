# coding: utf-8
"""
Base para desarrollo de modulos externos.
Para obtener el modulo/Funcion que se esta llamando:
     GetParams("module")

Para obtener las variables enviadas desde formulario/comando Rocketbot:
    var = GetParams(variable)
    Las "variable" se define en forms del archivo package.json

Para modificar la variable de Rocketbot:
    SetVar(Variable_Rocketbot, "dato")

Para obtener una variable de Rocketbot:
    var = GetVar(Variable_Rocketbot)

Para obtener la Opcion seleccionada:
    opcion = GetParams("option")


Para instalar librerias se debe ingresar por terminal a la carpeta "libs"

    pip install <package> -t .

"""
import os
import sys

base_path = tmp_global_obj["basepath"]
cur_path = base_path + 'modules' + os.sep + 'MicrosoftPowerpoint' + os.sep + 'libs' + os.sep
sys.path.append(cur_path)

# Import local libraries
import win32com.client

module = GetParams("module")
global powerpoint
global ms_pp


"""def alignments(WdParagraphAlignment):
    return ["Left", "Center", "Rigth", "Justify"][WdParagraphAlignment]


WdBuiltinStyle = {
    "paragraph": -1,
    "heading1": -2,
    "heading2": -3,
    "heading3": -4,
    "heading4": -5,
    "heading5": -6,
    "heading6": -7,
    "heading7": -8,
    "heading8": -9,
    "heading9": -10,
    "caption": -35,
    "bullet1": -49,
    "number1": -50,
    "bullet2": -55,
    "bullet3": -56,
    "bullet4": -57,
    "bullet5": -58,
    "number2": -59,
    "number3": -60,
    "number4": -61,
    "number5": -62,
    "title": -63,
    "subtitle": -75,
    "quote": -181,
    "intense_quote": -182,
    "book": -265
}"""

if module == "new":
    try:
        ms_pp = win32com.client.Dispatch("Powerpoint.Application")
        powerpoint = ms_pp.Presentations.Add()
    except Exception as e:
        print("\x1B[" + "31;40mError\u2193\x1B[" + "0m")
        PrintException()
        raise e

if module == "open":
    path = GetParams("path")

    try:
        ms_pp = win32com.client.DispatchEx("Powerpoint.Application")
        powerpoint = ms_pp.Presentations.Open(path.replace("/", os.sep))
        ms_pp.Visible = True
    except Exception as e:
        print("\x1B[" + "31;40mError\u2193\x1B[" + "0m")
        PrintException()
        raise e

if module == "save":

    path = GetParams("path")
    try:
        if path:
            powerpoint.SaveAs(path)
        else:
            powerpoint.SaveAs()
    except Exception as e:
        print("\x1B[" + "31;40mError\u2193\x1B[" + "0m")
        PrintException()
        raise e

if module == "to_pdf":
    path = GetParams("from")
    to = GetParams("to")
    ppFixedFormatTypePDF = 2
    try:
        if path:
            ms_pp = win32com.client.DispatchEx("Powerpoint.Application")
            powerpoint = ms_pp.Presentations.Open(path)
        powerpoint.SaveAs(FileName=to, FileFormat=32)
        powerpoint.Close()
        ms_pp.Quit()
    except Exception as e:
        print("\x1B[" + "31;40mError\x1B[" + "0m")
        PrintException()
        raise e

if module == "textbox":

    text = GetParams("text")
    align = GetParams("align")
    size = GetParams("size")
    bold = GetParams("bold")
    italic = GetParams("italic")
    underline = GetParams("underline")
    slide = GetParams("slide")
    pixelLeft = GetParams("pixelLeft")
    pixelTop = GetParams("pixelTop")
    pixelWidth = GetParams("pixelWidth")
    pixelHeight = GetParams("pixelHeight")
    try:
        currSlide = powerpoint.Slides(int(slide))
        textbox = currSlide.Shapes.AddTextBox(1, pixelLeft, pixelTop, pixelWidth, pixelHeight)
        range_ = textbox.TextFrame.TextRange


        range_.Text = text
        font = range_.Font

        size = float(size) if size else 12

        font.Size = size
        if bold == "True":
            boldInt = -1
        else: boldInt = 0
        font.Bold = boldInt
        if italic == "True":
            italicInt = -1
        else: italicInt = 0
        font.Italic = italicInt
        if underline == "True":
            underlineInt = -1
        else: underlineInt = 0
        font.Underline = underlineInt
        range_.ParagraphFormat.Alignment = int(align) if align else 0


    except Exception as e:
        PrintException()
        raise e

if module == "close":

    try:
        powerpoint.Close()
        ms_pp.Quit()
        powerpoint = None
        ms_pp = None
    except Exception as e:
        PrintException()
        raise e

if module == "new_slide":
    slide_number = GetParams("slide_number")
    try:
        powerpoint.Slides.Add(slide_number,12)
        print(powerpoint.Slides.count)
    except Exception as e:
        PrintException()
        raise e

if module == "add_pic":
    img_path = GetParams("img_path")
    slide_number = GetParams("slide_number")
    pixelLeft = GetParams("pixelLeft")
    pixelTop = GetParams("pixelTop")
    pixelWidth = GetParams("pixelWidth")
    pixelHeight = GetParams("pixelHeight")
    try:
        img_path = img_path.replace("/", os.sep)
        if pixelWidth and pixelHeight:
            img = powerpoint.Slides(int(slide_number)).Shapes.AddPicture(FileName=img_path, LinkToFile=False, SaveWithDocument=True, Left=pixelLeft, Top=pixelTop, Width=pixelWidth, Height=pixelHeight)
        elif pixelWidth:
            img = powerpoint.Slides(int(slide_number)).Shapes.AddPicture(FileName=img_path, LinkToFile=False,
                                                                   SaveWithDocument=True, Left=pixelLeft,
                                                                   Top=pixelTop, Width=pixelWidth)
        elif pixelHeight:
            img = powerpoint.Slides(int(slide_number)).Shapes.AddPicture(FileName=img_path, LinkToFile=False,
                                                                   SaveWithDocument=True, Left=pixelLeft,
                                                                   Top=pixelTop, Height=pixelHeight)
        else:
            img = powerpoint.Slides(int(slide_number)).Shapes.AddPicture(FileName=img_path, LinkToFile=False,
                                                                   SaveWithDocument=True, Left=pixelLeft,
                                                                   Top=pixelTop)
        print(img)
    except Exception as e:
        print("\x1B[" + "31;40mError\u2193\x1B[" + "0m")
        PrintException()
        raise e

if module == "search_text":
    try:
        text_search = GetParams("text_search")
        whichParagraph = GetParams("variable")
        paragraphList = []
        count = 1
        for paragraph in powerpoint.Paragraphs:
            range_ = paragraph.Range
            range_.Find.Text = text_search
            if range_.Find.Execute(Forward=True, MatchWholeWord=True):
                paragraphList.append(count)
            count += 1
        SetVar(whichParagraph, paragraphList)
        print(paragraphList)
    except Exception as e:
        print("\x1B[" + "31;40mError\u2193\x1B[" + "0m")
        PrintException()
        raise e

if module == "adjust":
    slide_number = GetParams("slide_number")
    numShape = GetParams("numShape")
    text = GetParams("text")
    rotation = GetParams("rotation")
    pixelLeft = GetParams("pixelLeft")
    pixelTop = GetParams("pixelTop")
    pixelWidth = GetParams("pixelWidth")
    pixelHeight = GetParams("pixelHeight")

    try:
        if text:
            powerpoint.Slides(int(slide_number)).Shapes(int(numShape)).TextFrame.TextRange.Text = text

        if rotation:
            powerpoint.Slides(int(slide_number)).Shapes(int(numShape)).Rotation = rotation

        if pixelLeft:
            powerpoint.Slides(int(slide_number)).Shapes(int(numShape)).Left(pixelLeft)

        if pixelTop:
            powerpoint.Slides(int(slide_number)).Shapes(int(numShape)).Top(pixelTop)

        if pixelWidth:
            powerpoint.Slides(int(slide_number)).Shapes(int(numShape)).Width(pixelWidth)

        if pixelHeight:
            powerpoint.Slides(int(slide_number)).Shapes(int(numShape)).Height(pixelHeight)

    except Exception as e:
        print("\x1B[" + "31;40mError\u2193\x1B[" + "0m")
        PrintException()
        raise e

if module == "add_shape":
    shape = GetParams("shape")
    slide_number = GetParams("slide_number")
    pixelLeft = GetParams("pixelLeft")
    pixelTop = GetParams("pixelTop")
    pixelWidth = GetParams("pixelWidth")
    pixelHeight = GetParams("pixelHeight")
    medialink = GetParams("medialink")
    text = GetParams("text")
    argv1 = {"Left": pixelLeft,
             "Top": pixelTop,
             "Width": pixelWidth,
             "Height": pixelHeight
             }
    argv2 = {key: value for key, value in argv1.items() if value}
    try:
        if shape == "label":
            powerpoint.Slides(int(slide_number)).Shapes.AddLabel(1,Left=pixelLeft,Top=pixelTop,Width=pixelWidth,Height=pixelHeight).TextFrame.TextRange.Text = text
        if shape == "title":
            if not powerpoint.Slides(int(slide_number)).Shapes.HasTitle:
                title = powerpoint.Slides(int(slide_number)).Shapes.AddTitle()
                title.TextFrame.TextRange.Text = text
        if shape == "media":
            medialink = medialink.replace("/", os.sep)
            powerpoint.Slides(int(slide_number)).Shapes.AddMediaObject2(FileName=medialink, LinkToFile=False, SaveWithDocument=True , **argv2)

    except Exception as e:
        print("\x1B[" + "31;40mError\u2193\x1B[" + "0m")
        PrintException()
        raise e

if module == "add_table":
    slide_number = GetParams("slide_number")
    numRows = GetParams("numRows")
    numCols = GetParams("numCols")
    pixelLeft = GetParams("pixelLeft")
    pixelTop = GetParams("pixelTop")
    pixelWidth = GetParams("pixelWidth")
    pixelHeight = GetParams("pixelHeight")
    print(pixelTop)
    print(type(pixelTop))
    try:
        argv1 = {"Left": pixelLeft,
                 "Top": pixelTop,
                 "Width": pixelWidth,
                 "Height": pixelHeight
                 }
        argv2 = {key:value for key, value in argv1.items() if value}
        powerpoint.Slides(int(slide_number)).Shapes.AddTable(numRows, numCols, **argv2)
    except Exception as e:
        print("\x1B[" + "31;40mError\u2193\x1B[" + "0m")
        PrintException()
        raise e

if module == "write_table":
    slide_number = GetParams("slide_number")
    shape_index = GetParams("shape_index")
    tblRow = GetParams("tblRow")
    tblCol = GetParams("tblCol")
    text = GetParams("text")
    try:
        powerpoint.Slides(int(slide_number)).Shapes(int(shape_index)).Table.Cell(tblRow,tblCol).Shape.TextFrame.TextRange.Text = text

    except Exception as e:
        print("\x1B[" + "31;40mError\u2193\x1B[" + "0m")
        PrintException()
        raise e
