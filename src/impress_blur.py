# -*- coding: utf-8 -*-
#!/usr/bin/env python

import uno
from com.sun.star.beans import PropertyValue
from com.sun.star.drawing.FillStyle import BITMAP

from PIL import Image, ImageFilter


try:
    from simple_dialogs import SelectBox, OptionBox, TextBox, NumberBox, DateBox, FolderPathBox, FilePathBox, MessageBox, ActionBox
except:
    from pythonpath.simple_dialogs import SelectBox, OptionBox, TextBox, NumberBox, DateBox, FolderPathBox, FilePathBox, MessageBox, ActionBox

#######################################################################
# start office with
# soffice "--accept=socket,host=127.0.0.1,port=2002,tcpNoDelay=1;urp;StarOffice.ComponentContext" --norestore
# then run script using python impress_blur.py to run extension 
# debugging using e.g. VS Code also possible that way
#######################################################################


# all the drawing components -> services
# sm.createInstanceWithContext("com.sun.star.drawing.GraphicExportFilter",ctx)
# BUT model.createInstance("com.sun.star.drawing.BitmapTable") for some reason
# https://api.libreoffice.org/docs/idl/ref/namespacecom_1_1sun_1_1star_1_1drawing.html



# -------------------------------------
# HELPERS FOR MRI AND  XRAY
# -------------------------------------

# Uncomment for MRI
# def mri(ctx, target):
#     mri = ctx.ServiceManager.createInstanceWithContext("mytools.Mri", ctx)
#     mri.inspect(target)

# Uncomment for Xray
# def xray(myObject):
#     try:
#         sm = uno.getComponentContext().ServiceManager
#         mspf = sm.createInstanceWithContext("com.sun.star.script.provider.MasterScriptProviderFactory", uno.getComponentContext())
#         scriptPro = mspf.createScriptProvider("")
#         xScript = scriptPro.getScript("vnd.sun.star.script:XrayTool._Main.Xray?language=Basic&location=application")
#         xScript.invoke((myObject,), (), ())
#         return
#     except:
#         raise _rtex("\nBasic library Xray is not installed", uno.getComponentContext())
# -------------------------------------------------------------------




def Run_impress_blur(*args):

    # java documentation kind of works: 
    # https://www.openoffice.org/api/docs/common/ref/com/sun/star/module-ix.html
    # https://www.openoffice.org/api/docs/common/ref/com/sun/star/presentation/module-ix.html

    # alternative: use XSCRIPTCONTEXT as context
    # https://www.openoffice.org/api/docs/common/ref/com/sun/star/script/provider/XScriptContext.html
    # context object

    try:
        ctx = remote_ctx                    # IDE, injected in __main__
    except:
        ctx = uno.getComponentContext()     # UI

    # get desktop
    # https://www.openoffice.org/api/docs/common/ref/com/sun/star/frame/XDesktop.html
    # A desktop is an environment for components which can be viewed in frames -> e.g. current file, model
    desktop = ctx.getByName("/singletons/com.sun.star.frame.theDesktop")

    # manager to load services, e.g. export service
    sm = ctx.getServiceManager()
    
    # https://www.openoffice.org/api/docs/common/ref/com/sun/star/frame/XModel.html
    # "represents a component which is created from an URL and arguments. 
    # It is a representation of a resource in the sense that it was created/loaded from the resource"
    # -> file + url
    model = desktop.getCurrentComponent()
    
    #alternative
    #model = XSCRIPTCONTEXT.getDocument()

    # always use URL for interaction with pyuno
    url = model.getURL()
    if url is None or url == "":
        m = MessageBox(message="Document doesn't have a filename. Save document first to use this plugin", title="No Filename")
        return

    # generate export file name
    print(f"file url = {url}")
    path = uno.fileUrlToSystemPath(url)
    print(f"file path = {path}")

    #get blurryness
    blurriness = NumberBox(message="Blurriness", title="Blurriness", default_value=5, min_=0, max_=20, decimals=0)
    if blurriness is None:
        return
    print(f" selected blurriness = {blurriness}")

    # get current slide

    # print(type(model.CurrentController.CurrentPage))
    # https://forum.openoffice.org/en/forum/viewtopic.php?t=105792
    # slide of type com.sun.star.drawing.XDrawPage (inspect with e.g. VS Code debugger)
    slide = model.CurrentController.CurrentPage

    # alternative: get slide by number (idk how to get numbers originally)
    # slide_num = 0
    # slide = model.DrawPages.getByIndex(slide_num)
    # number of slide object: slide.Number, see https://wiki.openoffice.org/wiki/Programming_OOoDraw_and_OOoImpress#Property_of_a_Slide
    print(f"select slide = {slide.Name} with number {slide.Number}")

    export_path = path + "_" + str(slide.Name) +  ".png"
    export_url = uno.systemPathToFileUrl(export_path)
    print(f"exporting to = {export_url} -> {export_path}")


    # image parameters
    filter_data = [PropertyValue(),PropertyValue()]
    
    filter_data[0].Name = "PixelWidth"
    filter_data[0].Value = 2560
    filter_data[1].Name = "PixelHeight"
    filter_data[1].Value = 1440

    # set export parameters
    prop_vals = [PropertyValue(),PropertyValue(),PropertyValue()]
    
    prop_vals[0].Name = "MediaType"
    prop_vals[0].Value = "image/png"

    prop_vals[1].Name = "URL"
    prop_vals[1].Value = export_url

    prop_vals[2].Name = "FilterData"
    prop_vals[2].Value = filter_data


    # export image
    xExporter = sm.createInstanceWithContext("com.sun.star.drawing.GraphicExportFilter",ctx)
    xExporter.setSourceDocument(slide)
    xExporter.filter(prop_vals)

    # blur it
    orig_im = Image.open(export_path)
    blur_image = orig_im.filter(ImageFilter.GaussianBlur(blurriness))
    blur_path = path + "_" + str(slide.Name) + "_blur_" + str(blurriness) + ".png"
    print(f"saving blur to = {blur_path}")
    blur_image.save(blur_path)


    # now add new slide with background
    # add slide: https://forum.openoffice.org/en/forum/viewtopic.php?t=55835
    # ehre an u.a. https://stackoverflow.com/questions/48398421/fill-the-background-of-a-presentation-file-with-pics-via-commandlines 
    # after current page: for some reason slide.Number starts at 1 and insertByIndex at 0 and insertByIndex
    # takes slide to be inserted after as index -> slide.Number-1
    new_slide = model.DrawPages.insertNewByIndex(slide.Number-1)

    #blur image path to url again
    blur_url = uno.systemPathToFileUrl(blur_path)

    # get bitmap manager for document -> all embeded images 
    # No idea why this service has to be created using the model
    bitmaps_manager = model.createInstance("com.sun.star.drawing.BitmapTable")
    
    # search for available name
    arb_number = 0
    blur_bitmap_name_base = "blur_" + slide.Name + "_"
    blur_bitmap_name = blur_bitmap_name_base + str(arb_number)
    while bitmaps_manager.hasByName(blur_bitmap_name):
        arb_number += 1
        blur_bitmap_name = blur_bitmap_name_base + str(arb_number)
    print(f"Inserting blurred bitmap with name = {blur_bitmap_name}")
    bitmaps_manager.insertByName(blur_bitmap_name,blur_url)
    
    background = model.createInstance("com.sun.star.drawing.Background")
    background.FillStyle = BITMAP
    background.FillBitmap = bitmaps_manager.getByName(blur_bitmap_name)
    new_slide.Background = background



g_exportedScripts = Run_impress_blur,


# -------------------------------------
# HELPER FOR AN IDE
# -------------------------------------

if __name__ == "__main__":
    """ Connect to LibreOffice proccess.
    1) Start the office in shell with command:
    soffice "--accept=socket,host=127.0.0.1,port=2002,tcpNoDelay=1;urp;StarOffice.ComponentContext" --norestore
    2) Run script
    """
    import os
    import sys

    sys.path.append(os.path.join(os.path.dirname(__file__), 'pythonpath'))

    local_ctx = uno.getComponentContext()
    resolver = local_ctx.ServiceManager.createInstance("com.sun.star.bridge.UnoUrlResolver")
    try:
        remote_ctx = resolver.resolve("uno:socket,"
                                        "host=localhost,"
                                        "port=2002,"
                                        "tcpNoDelay=1;"
                                        "urp;"
                                        "StarOffice.ComponentContext")
    except Exception as err:
        print(err)

    Run_impress_blur()