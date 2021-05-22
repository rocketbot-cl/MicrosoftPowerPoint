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
        ms_pp.Visible = True
    except Exception as e:
        print("\x1B[" + "31;40mError\u2193\x1B[" + "0m")
        PrintException()
        raise e

if module == "open":
    path = GetParams("path")

    try:
        ms_pp = win32com.client.DispatchEx("Powerpoint.Application")
        powerpoint = ms_pp.Presentations.Open(path)
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
        powerpoint.ExportAsFixedFormat(Path=to, FixedFormatType=ppFixedFormatTypePDF, IncludeDocProperties=True)
        powerpoint.Close()
        ms_pp.Quit()
    except Exception as e:
        print("\x1B[" + "31;40mError\x1B[" + "0m")
        PrintException()
        raise e

if module == "write":

    text = GetParams("text")
    type_ = GetParams("type")
    level = GetParams("level")
    align = GetParams("align")
    size = GetParams("size")
    bold = GetParams("bold")
    italic = GetParams("italic")
    underline = GetParams("underline")

    try:
        powerpoint.Paragraphs.Add()
        paragraph = powerpoint.Paragraphs.Last
        range_ = paragraph.Range
        range_.Text = text
        font = paragraph.Range.Font

        size = float(size) if size else 12

        font.Size = size
        font.Bold = bool(bold)
        font.Italic = bool(italic)
        font.Underline = bool(underline)

        paragraph.Alignment = int(align) if align else 0
        style = type_ + level
        if style in WdBuiltinStyle:
            paragraph.Style = WdBuiltinStyle[style]
        elif (type_ == "number" or type_ == "bullet") and int(level) > 5:
            level = 5
            style = type_ + str(level)
            paragraph.Style = WdBuiltinStyle[style]
        else:
            style = type_
            paragraph.Style = WdBuiltinStyle[style]
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
    try:
        powerpoint.Paragraphs.Add()
        paragraph = powerpoint.Paragraphs.Last
        paragraph.Range.InsertBreak()
    except Exception as e:
        PrintException()
        raise e

if module == "add_pic":
    img_path = GetParams("img_path")

    try:
        # Only work with \
        img_path = img_path.replace("/", os.sep)

        count = powerpoint.Paragraphs.Count  # Count number paragraphs
        if count > 1:
            powerpoint.Paragraphs.Add()

        paragraph = powerpoint.Paragraphs.Last
        img = paragraph.Range.InlineShapes.AddPicture(FileName=img_path, LinkToFile=False, SaveWithDocument=True)
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