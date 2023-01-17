"""powerpoint_win32.py implements App and Presentation for Powerpoint in Windows using COM interface.
"""

import enum
import traceback

# pip install pywin32
import pythoncom
import pywintypes
import win32api
import win32com.client

from process_exists import process_exists

from powerpoint_base import SlideCache, PresentationBase, PPTAppBase


pythoncom.CoInitialize()

ppLayoutBlank = 12  # Blank
ppLayoutChart = 8  # Chart
ppLayoutChartAndText = 6  # Chart and text
ppLayoutClipArtAndText = 10  # ClipArt and text
ppLayoutClipArtAndVerticalText = 26  # ClipArt and vertical text
ppLayoutComparison = 34  # Comparison
ppLayoutContentWithCaption = 35  # Content with caption
ppLayoutCustom = 32  # Custom
ppLayoutFourObjects = 24  # Four objects
ppLayoutLargeObject = 15  # Large object
ppLayoutMediaClipAndText = 18  # MediaClip and text
ppLayoutMixed = -2  # Mixed
ppLayoutObject = 16  # Object
ppLayoutObjectAndText = 14  # Object and text
ppLayoutObjectAndTwoObjects = 30  # Object and two objects
ppLayoutObjectOverText = 19  # Object over text
ppLayoutOrgchart = 7  # Organization chart
ppLayoutPictureWithCaption = 36  # Picture with caption
ppLayoutSectionHeader = 33  # Section header
ppLayoutTable = 4  # Table
ppLayoutText = 2  # Text
ppLayoutTextAndChart = 5  # Text and chart
ppLayoutTextAndClipArt = 9  # Text and ClipArt
ppLayoutTextAndMediaClip = 17  # Text and MediaClip
ppLayoutTextAndObject = 13  # Text and object
ppLayoutTextAndTwoObjects = 21  # Text and two objects
ppLayoutTextOverObject = 20  # Text over object
ppLayoutTitle = 1  # Title
ppLayoutTitleOnly = 11  # Title only
ppLayoutTwoColumnText = 3  # Two-column text
ppLayoutTwoObjects = 29  # Two objects
ppLayoutTwoObjectsAndObject = 31  # Two objects and object
ppLayoutTwoObjectsAndText = 22  # Two objects and text
ppLayoutTwoObjectsOverText = 23  # Two objects over text
ppLayoutVerticalText = 25  # Vertical text
ppLayoutVerticalTitleAndText = 27  # Vertical title and text
ppLayoutVerticalTitleAndTextOverChart = 28  # Vertical title and text over chart

# MsoShapeType: https://docs.microsoft.com/en-us/office/vba/api/office.msoshapetype
msoPlaceholder = 14
msoTextBox = 17

# PpPlaceholderType: https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppplaceholdertype
ppPlaceholderBitmap = 9  # Bitmap
ppPlaceholderBody = 2  # Body
ppPlaceholderCenterTitle = 3  # Center Title
ppPlaceholderChart = 8  # Chart
ppPlaceholderDate = 16  # Date
ppPlaceholderFooter = 15  # Footer
ppPlaceholderHeader = 14  # Header
ppPlaceholderMediaClip = 10  # Media Clip
ppPlaceholderMixed = -2  # Mixed
ppPlaceholderObject = 7  # Object
ppPlaceholderOrgChart = 11  # Organization Chart
ppPlaceholderPicture = 18  # Picture
ppPlaceholderSlideNumber = 13  # Slide Number
ppPlaceholderSubtitle = 4  # Subtitle
ppPlaceholderTable = 12  # Table
ppPlaceholderTitle = 1  # Title
ppPlaceholderVerticalBody = 6  # Vertical Body
ppPlaceholderVerticalObject = 17  # Vertical Object
ppPlaceholderVerticalTitle = 5  # Vertical Title

"""PpSaveAsFileType enumeration from
https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppsaveasfiletype
"""
ppSaveAsAddIn = 8
ppSaveAsAnimatedGIF = 40
ppSaveAsBMP = 19
ppSaveAsDefault = 11
ppSaveAsEMF = 23
ppSaveAsExternalConverter = 64000
ppSaveAsGIF = 16
ppSaveAsJPG = 17
ppSaveAsMetaFile = 15
ppSaveAsMP4 = 39
ppSaveAsOpenDocumentPresentation = 35
ppSaveAsOpenXMLAddin = 30
ppSaveAsOpenXMLPicturePresentation = 36
ppSaveAsOpenXMLPresentation = 24
ppSaveAsOpenXMLPresentationMacroEnabled = 25
ppSaveAsOpenXMLShow = 28
ppSaveAsOpenXMLShowMacroEnabled = 29
ppSaveAsOpenXMLTemplate = 26
ppSaveAsOpenXMLTemplateMacroEnabled = 27
ppSaveAsOpenXMLTheme = 31
ppSaveAsPDF = 32
ppSaveAsPNG = 18
ppSaveAsPresentation = 1
ppSaveAsRTF = 6
ppSaveAsShow = 7
ppSaveAsStrictOpenXMLPresentation = 38
ppSaveAsTemplate = 5
ppSaveAsTIF = 21
ppSaveAsWMV = 37
ppSaveAsXMLPresentation = 34
ppSaveAsXPS = 33


# https://docs.microsoft.com/en-us/dotnet/api/microsoft.office.core.msotristate?view=office-pia
class TriState(enum.Enum):
    msoFalse = 0
    msoMixed = -2
    msoTrue = -1


def TriStateToBool(value):
    return value != TriState.msoFalse


def BoolToTriState(value):
    return int(-1 if value else 0)


class PpParagraphAlignment(enum.IntEnum):
    ppAlignCenter = 2  # Center align
    ppAlignDistribute = 5  # Distribute
    ppAlignJustify = 4  # Justify
    ppAlignJustifyLow = 7  # Low justify
    ppAlignLeft = 1  # Left aligned
    ppAlignmentMixed = -2  # Mixed alignment
    ppAlignRight = 3  # Right-aligned
    ppAlignThaiDistribute = 6  # Thai distributed


class PpBaselineAlignment(enum.IntEnum):
    ppBaselineAlignBaseline = 1  # Aligned to the baseline.
    ppBaselineAlignCenter = 3  # Aligned to the center.
    ppBaselineAlignFarEast50 = 4  # Align FarEast50.
    ppBaselineAlignMixed = -2  # Mixed alignment.
    ppBaselineAlignTop = 2  # Aligned to the top.
    ppBaselineAlignAuto = 5  # https://github.com/DefinitelyTyped/DefinitelyTyped/blob/master/types/activex-powerpoint/index.d.ts


# https://github.com/DefinitelyTyped/DefinitelyTyped/blob/master/types/activex-powerpoint/index.d.ts
class PpExportMode(enum.IntEnum):
    ppClipRelativeToSlide = 2
    ppRelativeToSlide = 1
    ppScaleToFit = 3
    ppScaleXY = 4


# https://github.com/DefinitelyTyped/DefinitelyTyped/blob/a102789c764788888edc6a89542cd90f08fdce3d/types/activex-powerpoint/index.d.ts#L1014
class PpShapeFormat(enum.IntEnum):
    ppShapeFormatBMP = 3
    ppShapeFormatEMF = 5
    ppShapeFormatGIF = 0
    ppShapeFormatJPG = 1
    ppShapeFormatPNG = 2
    ppShapeFormatWMF = 4


class PpDirection(enum.IntEnum):
    ppDirectionLeftToRight = 1  # Left-to-right layout
    ppDirectionMixed = -2  # Mixed layout
    ppDirectionRightToLeft = 2  # Right-to-left layout


class PpViewType(enum.IntEnum):
    ppViewHandoutMaster = 4
    ppViewMasterThumbnails = 12
    ppViewNormal = 9
    ppViewNotesMaster = 5
    ppViewNotesPage = 3
    ppViewOutline = 6
    ppViewPrintPreview = 10
    ppViewSlide = 1
    ppViewSlideMaster = 2
    ppViewSlideSorter = 7
    ppViewThumbnails = 11
    ppViewTitleMaster = 8


class Font:
    def __init__(self, font):
        self.AutoRotateNumbers = TriStateToBool(font.AutoRotateNumbers)
        # a negative value automatically sets the Subscript property to True
        # a positive value automatically sets the Superscript property to True
        self.BaselineOffset = font.BaselineOffset
        self.Bold = TriState(font.Bold)
        # self.Color = font.Color
        self.Embeddable = TriStateToBool(font.Embeddable)
        self.Embedded = TriStateToBool(font.Embedded)
        self.Emboss = TriState(font.Emboss)
        self.Italic = TriState(font.Italic)
        self.Name = font.Name
        self.NameAscii = font.NameAscii
        self.NameComplexScript = font.NameComplexScript
        self.NameFarEast = font.NameFarEast
        self.NameOther = font.NameOther
        self.Shadow = TriState(font.Shadow)
        self.Size = font.Size
        # Setting the Subscript property to True automatically sets the BaselineOffset property to 0.3
        self.Subscript = TriState(font.Subscript)
        # Setting the Superscript property to True automatically sets the BaselineOffset property to - 0.25
        self.Superscript = TriState(font.Superscript)
        self.Underline = TriState(font.Underline)

    def __repr__(self):
        return self.__str__()

    def __str__(self):
        s = ""
        for key, value in self.__dict__.items():
            if s:
                s += ", "
            if key == "BaselineOffset":
                s += key + ": " + ("%g" % value)
            else:
                s += key + ": " + str(value)

        s = "Font(" + s + ")"

        return s


class ParagraphFormat:
    def __init__(self, pf):
        self.Alignment = PpParagraphAlignment(pf.Alignment)
        self.BaseLineAlignment = PpBaselineAlignment(pf.BaseLineAlignment)
        # self.Bullet = pf.Bullet
        self.FarEastLineBreakControl = TriStateToBool(pf.FarEastLineBreakControl)
        self.HangingPunctuation = TriStateToBool(pf.HangingPunctuation)
        self.LineRuleAfter = TriStateToBool(pf.LineRuleAfter)
        self.LineRuleBefore = TriStateToBool(pf.LineRuleBefore)
        self.LineRuleWithin = TriStateToBool(pf.LineRuleWithin)
        self.SpaceAfter = pf.SpaceAfter  # in points or lines
        self.SpaceBefore = pf.SpaceBefore  # in points or lines
        self.SpaceWithin = pf.SpaceWithin  # in line multiplier
        self.TextDirection = PpDirection(pf.TextDirection)
        self.WordWrap = pf.WordWrap

    def __repr__(self):
        return self.__str__()

    def __str__(self):
        s = ""
        for key, value in self.__dict__.items():
            if s:
                s += ", "
            if key == "SpaceAfter" or key == "SpaceBefore" or key == "SpaceWithin":
                s += key + ": " + ("%g" % value)
            else:
                s += key + ": " + str(value)
        s = "ParagraphFormat(" + s + ")"

        return s


def get_texts_in_shapes(shapes):
    texts = []
    for i in range(shapes.Count):
        shape = shapes.Item(i + 1)
        if not shape.HasTextFrame:
            continue

        text = shape.TextFrame.TextRange.Text
        texts.append(text)

    return texts


def find_text_in_shapes(shapes, text, match_case=True, whole_words=True):
    for i in range(shapes.Count):
        shape = shapes.Item(i + 1)
        if not shape.HasTextFrame:
            continue

        text_range = shape.TextFrame.TextRange
        found_text = text_range.Find(text, 0, BoolToTriState(match_case), BoolToTriState(whole_words))
        if found_text is not None:
            return found_text

    return None


def get_matching_textframe_info_in_shapes(shapes, text_dict):
    for i in range(shapes.Count):
        shape = shapes.Item(i + 1)
        if not shape.HasTextFrame:
            continue

        text_frame = shape.TextFrame
        text_range = text_frame.TextRange
        text = text_range.Text
        for from_text in text_dict:
            if from_text in text:
                found_range = text_range.Find(text)
                bbox = (shape.Left, shape.Top, shape.Width, shape.Height)
                mbox = (text_frame.MarginLeft, text_frame.MarginTop, text_frame.MarginRight, text_frame.MarginBottom)
                pf = ParagraphFormat(found_range.ParagraphFormat)
                font = Font(found_range.Font)
                return (from_text, bbox, mbox, pf, font)

    return None


# https://docs.microsoft.com/en-us/office/vba/api/powerpoint.textrange.replace
def replace_texts_in_shapes(shapes, text_dict):
    count = 0
    for i in range(shapes.Count):
        shape = shapes.Item(i + 1)
        if not shape.HasTextFrame:
            continue

        text_range = shape.TextFrame.TextRange
        text = text_range.Text
        for from_text, to_text in text_dict.items():
            if from_text in text:
                replaced_text = text_range.Replace(from_text, to_text)
                if replaced_text:
                    text = text.replace(from_text, to_text)
                    count = count + 1

    return count


"""
def dump_shapes(shapes):
    for i in range(shapes.Count):
        shape = shapes.Item(i+1)
        print('shape %d: %d' % (i, shape.Type))

        if shape.Type == msoPlaceholder:
            print(f'PlaceholderFormat.Type: {shape.PlaceholderFormat.Type}')

        if shape.HasTextFrame:
            text_range = shape.TextFrame.TextRange
            found_text = text_range.Text
            found_text = found_text.replace('\r', '\n')
            print('%s' % found_text)
"""


def is_text_placeholder(phType):
    """
    ppPlaceholderBody = 2 # Body
    ppPlaceholderCenterTitle = 3 # Center Title
    ppPlaceholderFooter = 15 # Footer
    ppPlaceholderHeader = 14 # Header
    ppPlaceholderSubtitle = 4 # Subtitle
    ppPlaceholderTitle = 1 # Title
    ppPlaceholderVerticalBody = 6 # Vertical Body
    ppPlaceholderVerticalTitle = 5 # Vertical Title
    """
    if ppPlaceholderTitle <= phType and phType <= ppPlaceholderVerticalBody:
        return True

    return False


def get_placeholder_text(shapes):
    text_dict = {}
    for i in range(shapes.Count):
        shape = shapes.Item(i + 1)

        if shape.Type == msoPlaceholder and shape.HasTextFrame:
            phType = shape.PlaceholderFormat.Type
            if is_text_placeholder(phType):
                text_range = shape.TextFrame.TextRange
                text = text_range.Text
                text_dict[phType] = text

    return text_dict


def copy_notes(source_range, dst_range):
    """copy_notes() fixes where SlideRange.Duplicate() does not copy NotesSlide."""
    # print('source_range')
    src_shapes = source_range.NotesPage.Shapes
    text_dict = get_placeholder_text(src_shapes)
    # dump_shapes(src_shapes)

    # print('dst_range')
    shapes = dst_range.NotesPage.Shapes
    # dump_shapes(shapes)
    for i in range(shapes.Count):
        shape = shapes.Item(i + 1)

        if shape.Type == msoPlaceholder and shape.HasTextFrame:
            phType = shape.PlaceholderFormat.Type
            if phType in text_dict:
                shape.TextFrame.TextRange.Text = text_dict[phType]

    # print('dst_range after change')
    # shapes = dst_range.NotesPage.Shapes
    # dump_shapes(shapes)


class Presentation(PresentationBase):
    """Presentation is a wrapper around PowerPoint.Presentation COM object.
    All the index used in the function is 0-based and converted to 1-based while calling COM functions.
    """

    def __init__(self, app, prs):
        super().__init__(app, prs)

    def close(self):
        if self.prs:
            self.prs.Close()
            self.prs = None

    def slide_count(self):
        return self.prs.Slides.Count

    def _fetch_slide_cache(self, slide_index):
        sc = SlideCache()

        slide = self.prs.Slides.Range(slide_index + 1)
        sc.id = slide.SlideID

        sc.slide_text = get_texts_in_shapes(slide.shapes)
        sc.notes_text = get_texts_in_shapes(slide.NotesPage.Shapes)

        sc.valid_cache = True

        return sc

    def slide_index_to_ID(self, var):
        if isinstance(var, int):
            slide = self.prs.Slides.Range(var + 1)
            return slide.SlideID
        elif isinstance(var, list):
            result = []
            for index in var:
                slide = self.prs.Slides.Range(index + 1)
                result.append(slide.SlideID)

            return result

    def slide_ID_to_index(self, var):
        if isinstance(var, int):
            slide = self.prs.Slides.FindBySlideID(var)
            return slide.SlideIndex - 1
        elif isinstance(var, list):
            result = []
            for sid in var:
                slide = self.prs.Slides.FindBySlideID(sid)
                result.append(slide.SlideIndex - 1)

            return result

    def _replace_texts_in_slide_shapes(self, slide_index, find_replace_dict):
        slide = self.prs.Slides.Range(slide_index + 1)
        shapes = slide.Shapes

        # tfinfo = get_matching_textframe_info_in_shapes(shapes, find_replace_dict)
        # if tfinfo:
        #    text = tfinfo[0]
        #    tfinfo = tfinfo[1:]
        #    print(f'text = {text}, {find_replace_dict[text]}')
        #    print(f'tfinfo = {tfinfo}')

        replace_texts_in_shapes(shapes, find_replace_dict)

    def duplicate_slides(self, source_location, insert_location=None, copy=1):

        source_count = 1
        if isinstance(source_location, int):
            source_location = source_location + 1
            source_count = 1
        elif isinstance(source_location, list):
            if len(source_location) == 1:
                source_location = source_location[0] + 1
                source_count = 1
            else:
                # make sure all elements are integer
                source_location = [int(i) + 1 for i in source_location]
        else:
            return 0

        # insert_location can change by copying slides.
        # So save the ID and use it later.
        if insert_location is None:
            insert_location = self.slide_count() - 1
        elif isinstance(insert_location, int):
            insert_location = insert_location + 1
        else:
            raise ValueError("Invalid insert location")

        added_count = 0
        source_range = self.prs.Slides.Range(source_location)

        # https://docs.microsoft.com/en-us/office/vba/api/powerpoint.sliderange.duplicate
        for _ in range(copy):
            insert_location1 = insert_location
            if insert_location > source_location:
                insert_location1 = insert_location1 + source_count

            dup_range = source_range.Duplicate()
            copy_notes(source_range, dup_range)
            dup_range.MoveTo(insert_location1)

            added_count = added_count + source_count

            self._slides_inserted(insert_location - 1, source_count)
            if source_location >= insert_location:
                source_location = source_location + source_count

            insert_location = insert_location + source_count

        return added_count

    def insert_blank_slide(self, slide_index):
        self.prs.Slides.Add(slide_index + 1, ppLayoutCustom)

        self._slides_inserted(slide_index, 1)

    def delete_slide(self, slide_index):
        self.prs.Slides.Range(slide_index + 1).Delete()

        self._slides_deleted(slide_index, 1)

    def copy_all_and_close(self):
        all_range = self.prs.Slides.Range()
        all_range.Copy()
        self.close()

    def _paste_keep_source_formatting(self, insert_location):
        # Activate any window associated with prs
        # w = self.prs.Windows.Item(1)
        # w.Activate()

        if self.slide_count() != 0:
            self.prs.Slides(insert_location).Select()

        # https://stackoverflow.com/questions/18457769/how-can-i-programatically-copy-and-paste-with-source-formatting-in-powerpoint-20
        powerpoint = self.prs.Application
        powerpoint.CommandBars.ExecuteMso("PasteSourceFormatting")

    def saveas(self, filename):
        """Save .pptx as files using Powerpoint.Application COM service."""
        try:
            self.prs.SaveAs(filename)
        except pywintypes.com_error as e:
            print("Error: %s\n%s" % (e, win32api.FormatMessage(e.args[2][5])))
            traceback.print_exc()

    def saveas_format(self, filename, image_type="png"):
        """Save .pptx as files using Powerpoint.Application COM service."""
        format_type = ppSaveAsPNG
        if image_type == "gif":
            format_type = ppSaveAsGIF
        elif image_type == "jpg":
            format_type = ppSaveAsJPG
        elif image_type == "png":
            format_type = ppSaveAsPNG
        elif image_type == "tif":
            format_type = ppSaveAsTIF

        self.prs.SaveAs(filename, format_type)

    def export_slide_as(self, slideno, filename, image_type="png"):
        """Save all shapes in slide as image files."""
        scale_width = 0
        scale_height = 0
        self.prs.Slides(slideno + 1).Export(filename, image_type, scale_width, scale_height)

    def export_shapes_as(self, slideno, filename, image_type="png"):
        """Save all shapes in slide as image file.
        It only exports the shapes not slide itself and the dimension may not what you expect,
        unless the whole selected shapes cover the slide.
        This method exports transparent images in good condition compared to the above saveas_format.

        https://stackoverflow.com/questions/5713676/ppt-to-png-with-transparent-background
        https://github.com/DefinitelyTyped/DefinitelyTyped/blob/a102789c764788888edc6a89542cd90f08fdce3d/types/activex-powerpoint/index.d.ts#L4234
        https://stackoverflow.com/questions/45299899/graphic-quality-in-powerpoint-graphics-export
        """
        window = self.app.powerpoint.ActiveWindow
        # window.Activate()
        window.ViewType = PpViewType.ppViewNormal
        window.View.GotoSlide(slideno + 1)
        slide = self.prs.Slides.Range(slideno + 1)
        slide.Shapes.SelectAll()

        format_type = PpShapeFormat.ppShapeFormatPNG
        if image_type == "gif":
            format_type = PpShapeFormat.ppShapeFormatGIF
        elif image_type == "png":
            format_type = PpShapeFormat.ppShapeFormatPNG

        # convert 1 inch (=72 point) to 112.5, 1 point to 1.562962963
        # This comes from manual export picture size.
        self.prs.PageSetup.SlideWidth
        scale_width = self.prs.PageSetup.SlideWidth * 1.563  # 1500
        scale_height = self.prs.PageSetup.SlideHeight * 1.563  # 844
        shape_range = window.Selection.ShapeRange
        shape_range.Export(filename, format_type, scale_width, scale_height, PpExportMode.ppRelativeToSlide)

    def close(self):
        try:
            self.prs.Close()
        except AttributeError:
            pass


class App(PPTAppBase):
    @staticmethod
    def is_running():
        return process_exists("powerpnt.exe")

    def __init__(self):
        self.powerpoint = win32com.client.Dispatch("Powerpoint.Application")
        self.powerpoint.Visible = 1

    def new_presentation(self):
        prs = self.powerpoint.Presentations.Add()
        if prs.Windows.Count == 0:
            prs.NewWindow()

        return Presentation(self, prs)

    def open_presentation(self, filename):
        prs = self.powerpoint.Presentations.Open(filename)
        if prs is None:
            return None

        return Presentation(self, prs)

    def quit(self, force=False, only_if_empty=True):
        call_quit = force
        if call_quit is False:
            if only_if_empty and self.powerpoint.Presentations.Count == 0:
                call_quit = True

        if call_quit is False:
            return

        self._quit()

    def _quit(self):
        try:
            self.powerpoint.Quit()
        except AttributeError:
            pass
