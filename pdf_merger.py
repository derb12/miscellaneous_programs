# -*- coding: utf-8 -*-
"""A GUI for selecting files to merge into one PDF file and save.

Can use the file formats natively supported by pymupdf (PDF, XPS, EPUB, HTML),
as well as image formats (JPG, TIF, PNG, SVG, GIF, BMP) and text files.

Also allows selecting the first page, last page, and rotation of each individual file
and allows previewing the merged file before saving.

@author: Donald Erb
Created on 2021-02-04
Copyright (C) 2020  Donald Erb

This program is free software: you can redistribute it and/or modify
it under the terms of the GNU Affero General Public License as published
by the Free Software Foundation, either version 3 of the License, or
(at your option) any later version.

This program is distributed in the hope that it will be useful,
but WITHOUT ANY WARRANTY; without even the implied warranty of
MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
GNU Affero General Public License for more details.

You should have received a copy of the GNU Affero General Public License
along with this program.  If not, see <https://www.gnu.org/licenses/>.
(See LICENSE.txt in this repository)

Requirements
------------
fitz>=1.18.4
    fitz is pymupdf; versions after 1.18.4 have snake_case rather than
    CamelCase or mixedCase.
wx>=4.0
    wx is wxpython; v4.0 is the first to work with python v3.

To install:

    pip install pymupdf>=1.18.4 wxPython>4.0

License
-------
This program is licensed under the GNU AGPL V3+ license (GNU AFFERO GPL).

History
-------
2021-10-16, Donald Erb
Added SAFE_SAVE attribute to allow saving with the default values,
in case the default saving options cause issues viewing the output
pdf.

Attributes
----------
CENTER_IMAGES_H : bool
    If True (default) and if FULL_PAGE_IMAGES is True, then will center
    images horizontally on the page.
CENTER_IMAGES_V : bool
    If True (default) and if FULL_PAGE_IMAGES is True, then will center
    images vertically on the page.
EXPAND_IMAGES : bool
    If True (default), will expand images to fit the size specified by
    PAGE_LAYOUT while retaining the original aspect ratio. If False, the
    pdf page with the image will directly fit the image's original size.
FONT : str
    The font name to use. If FONT_PATH is None, then it must be the name
    of a font covered by pymupdf. Default is 'helvetica'. Font is only
    used for converting text files to pdf.
FONT_PATH : str or None
    The file path to the font file to use, if not using a built-in font
    from pymupdf. Default is None, which means that FONT is a built-in font.
    Only used if converting text files to pdf.
FONT_SIZE : int
    The font size. Used for converting epub, html, and text files to pdf.
    Default is 11.
FULL_PAGE_IMAGES : bool
    If True, then images will retain their default size, or will be shrunked
    to fit the page size. This way, smaller images can keep their size, rather
    than being expanded to fill the page. Default is False. If True, images by
    default will be placed in the top-left corner, but can be centered using
    CENTER_IMAGES_H and CENTER_IMAGES_V.
PAGE_LAYOUT : str
    A string designating the paper size to use. Must be a valid input for
    fitz.PaperSize. Default is 'letter'. Append '-l' to change the page
    orientation.
USE_LANDSCAPE : bool
    If True, will append '-l' to PAGE_LAYOUT before passing it to fitz.PaperSize.
SAFE_SAVE : bool
    If True, will use the default options when saving, which do not compression
    or cleaning of the file. If False (default), will use the following options
    when saving:

        garbage=4, deflate=1

    Included since the more aggressive garbage collection used when SAFE_SAVE is False
    can sometimes cause issues when viewing the pdfs in Adobe (although the pdfs are fine
    when using other software like Chrome or Firefox to view the pdfs).

Notes
-----
If both EXPAND_IMAGES and FULL_PAGE_IMAGES are False when making a pdf page from an
image, then the pdf page will be made to fit the image, regardless of its size.

"""

import base64
import functools
import io
import os
from pathlib import Path
import re
import textwrap
import traceback

import fitz
import wx
import wx.grid


fitz.TOOLS.mupdf_display_errors(False)


CENTER_IMAGES_H = True
CENTER_IMAGES_V = True
EXPAND_IMAGES = True
FONT = 'Helvetica'
FONT_PATH = None
FONT_SIZE = 11
FULL_PAGE_IMAGES = False
PAGE_LAYOUT = 'letter'
USE_LANDSCAPE = False
SAFE_SAVE = False


if Path(__file__).parent.joinpath('logo.png').is_file():
    with Path(__file__).parent.joinpath('logo.png').open('rb') as fp:
        LOGO = base64.encodebytes(fp.read())
else:
    LOGO = None


# paper_sizes function replaced the paperSizes dictionary in a fitz version > 1.18.4
try:
    PAPER_SIZES = fitz.paper_sizes()
except AttributeError:
    PAPER_SIZES = fitz.paperSizes


class SettingsDialog(wx.Dialog):
    """
    Allows changing the default settings for page configuration.

    Parameters
    ----------
    parent : wx.Window, optional
        The parent widget for the dialog.
    **kwargs
        Any additional keyword arguments for initializing wx.Dialog.

    """

    def __init__(self, parent=None, **kwargs):
        super().__init__(parent, **kwargs)

        main_sizer = wx.BoxSizer(wx.HORIZONTAL)

        sizer_1 = wx.BoxSizer(wx.VERTICAL)
        main_sizer.Add(sizer_1, 1, wx.ALL | wx.EXPAND, 5)

        label_1 = wx.StaticText(
            self, label='Note: these settings are not used\nif directly using PDF files.',
            style=wx.ALIGN_CENTER_HORIZONTAL
        )
        sizer_1.Add(label_1, 0, wx.ALIGN_CENTER_HORIZONTAL | wx.BOTTOM, 15)

        self.expand_images = wx.RadioButton(
            self, label='Expand/shrink images to fill page', style=wx.RB_GROUP
        )
        self.expand_images.SetValue(EXPAND_IMAGES)
        sizer_1.Add(self.expand_images, 0, wx.BOTTOM | wx.TOP, 5)

        self.fit_pg_image = wx.RadioButton(self, label='Fit page to image')
        self.fit_pg_image.SetValue(not EXPAND_IMAGES and not FULL_PAGE_IMAGES)
        sizer_1.Add(self.fit_pg_image, 0, wx.BOTTOM | wx.TOP, 5)

        self.full_pg_image = wx.RadioButton(self, label='Keep image and page sizes')
        self.full_pg_image.SetValue(FULL_PAGE_IMAGES)
        sizer_1.Add(self.full_pg_image, 0, wx.BOTTOM | wx.TOP, 5)

        self.center_images_h = wx.CheckBox(self, label='Center images horizontally')
        sizer_1.Add(self.center_images_h, 0, wx.LEFT, 10)

        self.center_images_v = wx.CheckBox(self, label='Center images vertically')
        sizer_1.Add(self.center_images_v, 0, wx.LEFT | wx.BOTTOM, 10)

        if FULL_PAGE_IMAGES:
            self.center_images_h.SetValue(CENTER_IMAGES_H)
            self.center_images_v.SetValue(CENTER_IMAGES_V)
        else:
            self.center_images_h.SetValue(False)
            self.center_images_v.SetValue(False)
            self.center_images_h.Enable(False)
            self.center_images_v.Enable(False)

        # filler
        sizer_1.Add((20, 20))

        sizer_6 = wx.BoxSizer(wx.HORIZONTAL)
        sizer_1.Add(sizer_6, 1, wx.BOTTOM | wx.EXPAND | wx.TOP, 5)
        label_4 = wx.StaticText(self, label='Page Layout')
        sizer_6.Add(label_4, 0, wx.ALIGN_CENTER_VERTICAL | wx.RIGHT, 10)

        paper_sizes = list(PAPER_SIZES.keys())
        self.page_layout = wx.Choice(self, wx.ID_ANY, choices=paper_sizes)
        if PAGE_LAYOUT in PAPER_SIZES:
            selection = paper_sizes.index(PAGE_LAYOUT)
        else:
            selection = paper_sizes.index('letter')
        self.page_layout.SetSelection(selection)
        sizer_6.Add(self.page_layout, 0, wx.ALIGN_CENTER_VERTICAL, 0)

        self.landscape = wx.CheckBox(self, label='Use landscape (makes width > height)')
        self.landscape.SetValue(USE_LANDSCAPE)
        sizer_1.Add(self.landscape, 0, wx.BOTTOM | wx.TOP, 5)

        self.safe_save = wx.CheckBox(self, label='Safe Save? Check if the output files have issues')
        self.safe_save.SetValue(SAFE_SAVE)
        sizer_1.Add(self.safe_save, 0, wx.BOTTOM | wx.TOP, 5)

        # filler
        sizer_1.Add((20, 20))

        sizer_4 = wx.BoxSizer(wx.HORIZONTAL)
        sizer_1.Add(sizer_4, 1, wx.BOTTOM | wx.EXPAND | wx.TOP, 5)

        label_2 = wx.StaticText(self, label='Font')
        sizer_4.Add(label_2, 0, wx.ALIGN_CENTER_VERTICAL | wx.RIGHT, 10)

        self.font = wx.Choice(self, choices=fitz.Base14_fontnames)
        if FONT in fitz.Base14_fontnames:
            selection = fitz.Base14_fontnames.index(FONT)
        else:
            selection = fitz.Base14_fontnames.index('Helvetica')
        self.font.SetSelection(selection)
        sizer_4.Add(self.font, 0, wx.ALIGN_CENTER_VERTICAL, 0)

        sizer_5 = wx.BoxSizer(wx.HORIZONTAL)
        sizer_1.Add(sizer_5, 1, wx.BOTTOM | wx.EXPAND | wx.TOP, 5)

        label_3 = wx.StaticText(self, label='Font Size')
        sizer_5.Add(label_3, 0, wx.ALIGN_CENTER_VERTICAL | wx.RIGHT, 10)

        self.font_size = wx.SpinCtrl(self, value=str(int(float(FONT_SIZE))), min=1, max=256)
        sizer_5.Add(self.font_size, 0, wx.ALIGN_CENTER_VERTICAL, 0)

        # filler
        sizer_1.Add((20, 30))

        sizer_3 = wx.StdDialogButtonSizer()
        sizer_1.Add(sizer_3, 0, wx.ALIGN_RIGHT | wx.ALL, 5)
        self.button_OK = wx.Button(self, wx.ID_OK)
        self.button_OK.SetDefault()
        sizer_3.AddButton(self.button_OK)
        self.button_CANCEL = wx.Button(self, wx.ID_CANCEL)
        sizer_3.AddButton(self.button_CANCEL)
        sizer_3.Realize()

        self.SetAffirmativeId(self.button_OK.GetId())
        self.SetEscapeId(self.button_CANCEL.GetId())

        self.SetSizer(main_sizer)
        main_sizer.Fit(self)

        self.expand_images.Bind(wx.EVT_RADIOBUTTON, self.on_radio)
        self.full_pg_image.Bind(wx.EVT_RADIOBUTTON, self.on_radio)
        self.fit_pg_image.Bind(wx.EVT_RADIOBUTTON, self.on_radio)

    def on_radio(self, event):
        """Enables or disables image options depending on self.full_pg_image."""
        checked = self.full_pg_image.GetValue()
        self.center_images_h.SetValue(checked)
        self.center_images_v.SetValue(checked)
        self.center_images_h.Enable(checked)
        self.center_images_v.Enable(checked)

    def set_options(self):
        """Overrides the global config variables."""
        global CENTER_IMAGES_H
        global CENTER_IMAGES_V
        global EXPAND_IMAGES
        global FONT
        global FONT_SIZE
        global FULL_PAGE_IMAGES
        global PAGE_LAYOUT
        global USE_LANDSCAPE
        global SAFE_SAVE

        CENTER_IMAGES_H = self.center_images_h.GetValue()
        CENTER_IMAGES_V = self.center_images_v.GetValue()
        EXPAND_IMAGES = self.expand_images.GetValue()
        FONT = self.font.GetStringSelection()
        FONT_SIZE = self.font_size.GetValue()
        FULL_PAGE_IMAGES = self.full_pg_image.GetValue()
        PAGE_LAYOUT = self.page_layout.GetStringSelection()
        USE_LANDSCAPE = self.landscape.GetValue()
        SAFE_SAVE = self.safe_save.GetValue()


class PagesGrid(wx.grid.Grid):
    """
    Contains information about the pages and rotation of pdf files.

    Parameters
    ----------
    parent : wx.Window
        The parent widget.
    **kwargs
        Any additional keyword arguments for initializing wx.Grid.
    """

    def __init__(self, parent, **kwargs):
        super().__init__(parent, **kwargs)

        self.CreateGrid(0, 5)
        self.EnableDragRowSize(False)
        self.HideCol(4)  # column 4 is just to hold the total number of pages
        self.SetSelectionMode(wx.grid.Grid.SelectRows)

        attr = wx.grid.GridCellAttr()
        attr.SetReadOnly(True)
        self.SetColAttr(0, attr)
        attr.IncRef()

        attr = wx.grid.GridCellAttr()
        attr.SetAlignment(wx.ALIGN_CENTER, -1)
        for col in range(1, 4):
            self.SetColAttr(col, attr)
            attr.IncRef()

        attr = wx.grid.GridCellAttr()
        attr.SetAlignment(wx.ALIGN_CENTER, -1)
        attr.SetReadOnly(True)
        self.SetColAttr(4, attr)

        self.SetColLabelValue(0, 'File')
        self.SetColSize(0, 400)
        self.SetColLabelValue(1, 'First Page')
        self.SetColLabelValue(2, 'Last Page')
        self.SetColLabelValue(3, 'Rotation')
        self.SetColLabelValue(4, 'Total Pages')

    def add_row(self, file_path, row_index=None):
        """
        Appends a row to the grid with information after opening the given file.

        Parameters
        ----------
        file_path : str or os.Pathlike
            The path of the pdf file to open and read.
        row_index : int, optional
            The index to insert a new row. Default is None, which will append
            the row to the grid.

        """
        try:
            pdf = get_pdf(file_path)
            if pdf.needs_pass:  # password protected
                raise ValueError('File is encrypted and cannot be processed')
            pages = len(pdf)
        except Exception:
            with wx.MessageDialog(
                self,
                (f'Problem opening {file_path}\n\nError:\n    {traceback.format_exc()}'),
            ) as dlg:
                dlg.ShowModal()
            return
        finally:
            try:
                pdf.close()
            except Exception:
                pass

        if row_index is None:
            row = self.GetNumberRows()
            self.AppendRows(1)
        else:
            row = row_index
            self.InsertRows(row)
        self.create_row(row, file_path, pages)

    def create_row(self, row, file_path, total_pages, first_page='1',
                   last_page=None, rotation=None):
        """
        Adds data about a pdf file to a row in the grid.

        Note that this function does not handle the actual creation of the row, so
        that must be done before calling this method.

        Parameters
        ----------
        row : int
            The index of the row to add information to.
        file_path : str or os.Pathlike
            The file path of the pdf file.
        total_pages : int
            The total number of pages in the pdf file.
        first_page : str or int, optional
            The first page of the pdf to use; by default '1'.
        last_page : str or int, optional
            The last page of the pdf to use. If None, will be set to the last page.
        rotation : {0, 90, -90, 180}, optional
            The integer rotation (in degrees) of the pdf file.

        """
        pages = [str(page + 1) for page in range(total_pages)]
        first_pg = first_page if str(first_page) in pages else pages[0]
        last_pg = last_page if last_page is not None else pages[-1]
        rotations = {0: '0°', -90: '-90° (left)', 90: '90° (right)', 180: '180°'}
        if rotation is None:
            start_rotation = rotations[0]
        elif int(rotation) in rotations.keys():
            start_rotation = rotations[int(rotation)]
        elif str(rotation) in rotations.values():
            start_rotation = str(rotation)
        else:
            start_rotation = rotations[0]

        self.SetCellValue(row, 0, str(file_path))
        self.SetCellEditor(row, 1, wx.grid.GridCellChoiceEditor(pages, False))
        self.SetCellValue(row, 1, str(first_pg))
        self.SetCellEditor(row, 2, wx.grid.GridCellChoiceEditor(pages, False))
        self.SetCellValue(row, 2, str(last_pg))
        self.SetCellEditor(row, 3, wx.grid.GridCellChoiceEditor(list(rotations.values()), False))
        self.SetCellValue(row, 3, start_rotation)
        self.SetCellValue(row, 4, str(total_pages))

    def get_values(self):
        """
        Returns the relevant info for each pdf file in the grid.

        Returns
        -------
        data : list(list(str, int, int, int))
            A list of lists. Each internal list should have four items, telling
            the file path for each pdf to merge, the first page to use, the last
            page to use, and the rotation. Each individual entry is as follows:

                file_path: str
                    The collection of files to merge into a single document and saved.
                first_pg: int
                    The first page to use. 0-based.
                last_pg: int
                    The last page to use. 0-based. If less than the first_pg, then the
                    page order will be reversed.
                rotations : {0, 90, -90, 180}
                    The integer rotation to apply to the document. Note that -90 will
                    rotate the document left (counter-clockwise).
        """
        data = []
        for row in range(self.GetNumberRows()):
            data.append([
                self.GetCellValue(row, 0),
                int(self.GetCellValue(row, 1)) - 1,
                int(self.GetCellValue(row, 2)) - 1,
                int(self.GetCellValue(row, 3).split('°')[0]),
            ])

        return data


class PDFMerger(wx.Frame):
    """
    A frame for selecting PDF files to merge.

    Also allows selecting the first page, last page, and rotation of each
    individual PDF and viewing a preview of the output file.

    Parameters
    ----------
    parent : wx.Window, optional
        The parent widget for this frame. Default is None.
    **kwargs
        Any additional keyword arguments for initializing wx.Frame.

    """

    def __init__(self, parent=None, **kwargs):
        super().__init__(parent, **kwargs)
        self.SetSize(self.FromDIP((900, 500)))
        self.preview = None
        if LOGO is not None:
            self.SetIcon(wx.Icon(wx.Image(io.BytesIO(base64.b64decode(LOGO))).ConvertToBitmap()))

        self.menubar = wx.MenuBar()
        self.options_menu = wx.Menu('Set Options')
        self.menubar.Append(self.options_menu, "Options")
        self.SetMenuBar(self.menubar)

        self.panel = wx.Panel(self)
        sizer = wx.BoxSizer(wx.VERTICAL)

        sizer_1 = wx.BoxSizer(wx.HORIZONTAL)
        sizer.Add(sizer_1, 0, wx.EXPAND, 0)

        grid_sizer_1 = wx.GridSizer(1, 2, 0, 20)
        sizer_1.Add(grid_sizer_1, 0, wx.ALIGN_CENTER_VERTICAL | wx.ALL, 10)

        self.add_btn = wx.Button(self.panel, label='Add Files')
        grid_sizer_1.Add(self.add_btn, 0, wx.EXPAND, 0)

        self.remove_btn = wx.Button(self.panel, label='Remove Files')
        grid_sizer_1.Add(self.remove_btn, 0, wx.EXPAND, 0)

        label_1 = wx.StaticText(
            self.panel,
            label=('Use "Add Files" to add multiple files, "Remove Files" to'
                   '\nremove selected files, and ▲ or ▼ to reorder files.'),
            style=wx.ALIGN_CENTER_HORIZONTAL
        )
        sizer_1.Add(label_1, 1, wx.BOTTOM | wx.TOP, 5)

        sizer_2 = wx.BoxSizer(wx.HORIZONTAL)
        sizer.Add(sizer_2, 1, wx.ALL | wx.EXPAND, 8)

        sizer_5 = wx.BoxSizer(wx.VERTICAL)
        sizer_2.Add(sizer_5, 1, wx.EXPAND, 0)

        self.grid = PagesGrid(self.panel, size=(1, 1))
        sizer_5.Add(self.grid, 1, wx.ALL | wx.EXPAND, 0)

        sizer_3 = wx.BoxSizer(wx.HORIZONTAL)
        sizer_5.Add(sizer_3, 0, wx.BOTTOM | wx.EXPAND | wx.TOP, 20)

        self.output_file = wx.TextCtrl(self.panel, style=wx.TE_READONLY)
        sizer_3.Add(self.output_file, 1, wx.EXPAND | wx.RIGHT, 5)

        self.saveas_btn = wx.Button(self.panel, wx.ID_SAVEAS)
        sizer_3.Add(self.saveas_btn, 0, wx.ALIGN_CENTER_VERTICAL, 30)

        grid_sizer_2 = wx.GridSizer(2, 1, 0, 0)
        sizer_2.Add(grid_sizer_2, 0, wx.ALIGN_CENTER_VERTICAL | wx.ALL, 10)

        self.up_btn = wx.Button(self.panel, wx.ID_ANY, '▲', style=wx.BU_EXACTFIT)
        grid_sizer_2.Add(self.up_btn, 0, wx.EXPAND, 0)

        self.down_btn = wx.Button(self.panel, wx.ID_ANY, '▼', style=wx.BU_EXACTFIT)
        grid_sizer_2.Add(self.down_btn, 0, wx.EXPAND, 0)

        sizer_4 = wx.BoxSizer(wx.HORIZONTAL)
        sizer.Add(sizer_4, 0, wx.ALL, 10)

        self.save_btn = wx.Button(self.panel, wx.ID_SAVE)
        sizer_4.Add(self.save_btn, 0, wx.RIGHT, 10)

        self.preview_btn = wx.Button(self.panel, label='Preview')
        sizer_4.Add(self.preview_btn, 0, 0, 0)

        self.panel.SetSizer(sizer)
        self.Layout()

        self.add_btn.Bind(wx.EVT_BUTTON, self.on_add)
        self.remove_btn.Bind(wx.EVT_BUTTON, self.on_remove)
        self.up_btn.Bind(wx.EVT_BUTTON, self.move_up)
        self.down_btn.Bind(wx.EVT_BUTTON, self.move_down)
        self.saveas_btn.Bind(wx.EVT_BUTTON, self.on_saveas)
        self.save_btn.Bind(wx.EVT_BUTTON, self.on_save)
        self.preview_btn.Bind(wx.EVT_BUTTON, self.on_preview)
        self.Bind(wx.EVT_MENU, self.set_options, self.menubar)

    def set_options(self, event):
        """
        Launches dialog to override the global settings.

        If the dialog is confirmed, then the grid data will be reset for all file
        types except for pdf and xps files so that the page information can be
        updated. Image files will keep their rotations as well.

        """
        reset_grid = False
        with SettingsDialog(self, title='Options') as dialog:
            if dialog.ShowModal() == wx.ID_OK:
                dialog.set_options()
                reset_grid = True

        if reset_grid:
            grid_data = self.grid.get_values()
            rotations = {0: '0°', -90: '-90° (left)', 90: '90° (right)', 180: '180°'}
            for row, row_data in enumerate(grid_data):
                suffix = Path(row_data[0]).suffix.lower()
                if re.search('.*xps|pdf', suffix) is None:
                    self.grid.DeleteRows(row)
                    self.grid.add_row(row_data[0], row)
                    if is_image(suffix):
                        self.grid.SetCellValue(row, 3, rotations[row_data[3]])
        event.Skip()

    def on_add(self, event):
        """Launches file dialog to select pdf files and adds them to the grid."""
        paths = []
        with wx.FileDialog(
            self, 'Select files',
            wildcard=(
                'PDF Files|*.pdf|Image Files|*.jp*;*.png;*.tif*;*.svg;*.gif;*.bmp|'
                'XPS Files|*.*xps|HTML Files|*.htm*|EPUB Files|*.epub|All Files|*.*'
            ),
            style=wx.FD_OPEN | wx.FD_CHANGE_DIR | wx.FD_FILE_MUST_EXIST | wx.FD_MULTIPLE
        ) as dialog:
            if dialog.ShowModal() == wx.ID_OK:
                # maybe not necessary, but ensures paths are correct for all os
                paths = [str(Path(path)) for path in dialog.GetPaths()]
        for path in paths:
            self.grid.add_row(path)

    def on_remove(self, event):
        """Removes selected files from the grid."""
        rows = self.grid.GetSelectedRows()
        for row in reversed(rows):
            self.grid.DeleteRows(row)
        self.grid.ClearSelection()

    def _move(self, move_down=False):
        """Moves all selected grid rows up or down, if possible."""
        rows = self.grid.GetSelectedRows()
        if not rows:
            return
        self.grid.ClearSelection()
        failed_moves = []
        if move_down:
            new_row = lambda x: x + 2
            old_row = lambda x: x
            select_row = lambda x: x + 1
            failed_row = lambda x: x + 1
            rows = sorted(rows, reverse=True)
            if rows[0] == self.grid.GetNumberRows() - 1:
                failed_moves.append(rows.pop(0) + 1)
        else:
            new_row = lambda x: x - 1
            old_row = lambda x: x + 1
            select_row = lambda x: x - 1
            failed_row = lambda x: x
            if rows[0] == 0:
                failed_moves.append(rows.pop(0))

        grid_data = self.grid.get_values()
        for row in rows:
            if new_row(row) in failed_moves:
                failed_moves.append(failed_row(row))
                continue

            self.grid.InsertRows(new_row(row))
            file_path, first_pg, last_pg, rotation = grid_data[row]
            self.grid.create_row(
                new_row(row), file_path, int(self.grid.GetCellValue(old_row(row), 4)),
                first_pg + 1, last_pg + 1, rotation
            )
            self.grid.DeleteRows(old_row(row))
            self.grid.SelectRow(select_row(row), True)

    def move_up(self, event):
        """Moves all selected grid rows up, if possible."""
        self._move(False)
        event.Skip()

    def move_down(self, event):
        """Moves all selected grid rows down, if possible."""
        self._move(True)
        event.Skip()

    def on_saveas(self, event):
        """Selects the output file name and directory."""
        with wx.FileDialog(
            self, 'Save As...',
            wildcard='PDF Files (*.pdf)|*.pdf',
            style=wx.FD_SAVE | wx.FD_OVERWRITE_PROMPT
        ) as dialog:
            if self.output_file.GetValue():
                path = Path(self.output_file.GetValue())
                dialog.SetFilename(path.name)
                dialog.SetDirectory(str(path.parent))

            if dialog.ShowModal() == wx.ID_OK:
                self.output_file.SetValue(str(Path(dialog.GetPath())))

    def on_save(self, event):
        """
        Merges the selected pdfs and saves.

        The frame is closed if the merged pdf is saved successfully.

        """
        event.Skip()
        error_msg = ''
        output_path = self.output_file.GetValue()
        grid_data = self.grid.get_values()
        if not output_path:
            error_msg = 'Need to select the output file name.'
        elif not grid_data:
            error_msg = 'No PDF files to merge.'

        if error_msg:
            with wx.MessageDialog(self, error_msg, 'Error') as dlg:
                dlg.ShowModal()
            return

        try:
            output_pdf = merge_pdfs(grid_data, True)
            if len(output_pdf) == 0:
                raise ValueError('The output pdf has no pages.')
            # note: garbate > 2 will merge the same objects, which can cause issues viewing
            # the pdf with Adobe (although the pdf can still be viewed with other software)
            if SAFE_SAVE:
                save_options = {}
            else:
                save_options = {'garbage': 4, 'deflate': 1}
            output_pdf.save(output_path, **save_options)
        except Exception:
            with wx.MessageDialog(
                self, f'Could not save file\n\n    {traceback.format_exc()}',
                'Error Saving'
            ) as dlg:
                dlg.ShowModal()
        else:
            with wx.MessageDialog(
                self, f'File successfully created\n\nFile at:\n    {output_path}\n',
                'Save Successful'
            ) as dlg:
                dlg.ShowModal()
            self.output_file.SetValue('')
        finally:
            try:
                output_pdf.close()
            except Exception:
                pass

    def on_preview(self, event):
        """Launches the pdf viewer to show a preview."""
        grid_values = self.grid.get_values()
        if self.preview or not grid_values:
            return

        try:
            temp_pdf = merge_pdfs(grid_values)
        except Exception:
            with wx.MessageDialog(
                self, f'Could not make preview\n\n    {traceback.format_exc()}',
                'Error with preview'
            ) as dlg:
                dlg.ShowModal()
            temp_pdf = None
            return

        self.preview = PDFViewer(self, temp_pdf, title='PDF Preview')
        self.preview.Show()
        try:
            temp_pdf.close()
        except ValueError:
            pass


def merge_pdfs(grid_data, finalize=False):
    """
    Merges all of the input PDF files into a single PDF and saves.

    Parameters
    ----------
    grid_data : list(list(str, int, int, int))
        A list of lists. Each internal list should have four items, telling
        the file path for each pdf to merge, the first page to use, the last
        page to use, and the rotation. Each individual entry is as follows:

            file_path: str
                The collection of files to merge into a single document and saved.
            first_pg: int
                The first page to use. 0-based.
            last_pg: int
                The last page to use. 0-based. If less than the first_pg, then the
                page order will be reversed.
            rotations : {0, 90, -90, 180}
                The integer rotation to apply to the document. Note that -90 will
                rotate the document left (counter-clockwise).
    finalize : bool, optional
        If False (default), will ignore links when merging the pdfs and will not
        create a table of contents. If True, will keep links and create the full
        merged table of contents.

    """
    current_pg = 0
    output_file = fitz.Document()
    total_toc = []  # collects the bookmarks from all of the files
    for file_path, first_pg, last_pg, rotation in grid_data:
        path = Path(file_path)

        with get_pdf(path, finalize) as temp_file:
            pages = len(temp_file)

            # ensures pages are within the document
            first_pg = min(max(0, first_pg), pages - 1)
            if last_pg != -1:
                last_pg = min(max(0, last_pg), pages - 1)
            else:
                last_pg = pages - 1

            # only add unencrypted files
            if not temp_file.needs_pass:
                output_file.insert_pdf(
                    temp_file, from_page=first_pg, to_page=last_pg,
                    rotate=rotation, links=finalize, annots=True
                )
            else:
                print(
                    f'\nThe following file is encrypted and cannot be processed:\n\n    {str(path)}'
                )
                continue

            if not finalize:
                continue  # skip creating the table of contents

            # get file's table of contents
            toc = temp_file.get_toc(simple=False)
            if first_pg > last_pg:
                increment = -1
                toc = toc[::-1]
            else:
                increment = 1

            pg_range = list(range(first_pg, last_pg + increment, increment))

            # set starting bookmark level to 1
            last_lvl = 1
            for link in toc:
                lnk_type = link[3]["kind"]

                if lnk_type == fitz.LINK_NAMED:
                    # skip named links since pymupdf cannot process them
                    continue
                elif lnk_type == fitz.LINK_GOTO:
                    if link[2] - 1 not in pg_range:
                        # skip the bookmark if it's not within the page range
                        continue
                    else:
                        page_num = pg_range.index(link[2] - 1) + current_pg + 1

                        # fix bookmark levels left by filler bookmarks
                        while (link[0] > last_lvl + 1):
                            total_toc.append([last_lvl + 1, "<>", page_num, link[3]])
                            last_lvl += 1

                        last_lvl = link[0]
                        link[2] = page_num
                total_toc.append(link)

            current_pg += len(pg_range)

    if total_toc:
        output_file.set_toc(total_toc)

    return output_file


def get_page_layout():
    """Returns the selected page layout."""
    page_layout = PAGE_LAYOUT
    if USE_LANDSCAPE and not page_layout.endswith('-l'):
        page_layout += '-l'
    return page_layout


def get_pdf(file_name, finalize=False):
    """
    Generates a pymupdf Document based on the file extension of the file.

    Parameters
    ----------
    file_name : str or os.Pathlike
        The file path for the document.
    finalize : bool, optional
        If False (default), will not copy table of contents for xps, html, or
        epub files. If True, will copy the table of contents.

    Returns
    -------
    fitz.Document
        The file converted to a pymupdf Document using the appropriate conversions.

    """
    path = Path(file_name)
    suffix = path.suffix.lower()
    if suffix.startswith('.'):
        suffix = suffix[1:]

    if suffix == 'pdf':
        return fitz.Document(file_name)
    elif re.search('.*xps|epub|htm.*', suffix) is not None:
        return document_to_pdf(file_name, finalize)
    elif is_image(suffix):
        return image_to_pdf(file_name)
    else:
        return text_to_pdf(file_name)


def document_to_pdf(file_name, finalize=False):
    """
    Converts epub, html, and xps files to pdf while retaining links and table of contents.

    Parameters
    ----------
    file_name : str or os.Pathlike
        The file path for the document.
    finalize : bool, optional
        If False (default), will not copy table of contents for xps, html, or
        epub files. If True, will copy the table of contents.

    Returns
    -------
    pdf : fitz.Document
        The file converted to a pymupdf Document using the appropriate conversions.

    Notes
    -----
    htm* and epub files are sized according to PAGE_LAYOUT and FONT_SIZE since
    their size can be variable.

    """
    path = Path(file_name)
    width, height = fitz.PaperSize(get_page_layout())
    with fitz.Document(str(path), width=width, height=height, fontsize=FONT_SIZE) as original:
        pdf = fitz.Document(filetype='pdf', stream=original.convert_to_pdf())
        if finalize:
            pdf.set_toc(original.get_toc())

    return pdf


def text_to_pdf(file_name):
    """
    Converts text files to pdf.

    Parameters
    ----------
    file_name : str or os.Pathlike
        The file path for the text document.

    Returns
    -------
    pdf : fitz.Document
        The file converted to a pymupdf Document.

    Notes
    -----
    Uses the default pymupdf settings to determine the maximum lines per page and
    characters per line (although the max lines per page is slightly different...?).

    Works okay when using Helvetica and size 11 font, but is not guaranteed to work
    for all cases. Best to just use Word to convert text to pdf, if available.

    """
    width, height = fitz.PaperSize(get_page_layout())
    # 108 is the header + footer spacing (72 is the botton of the first inserted
    # row, and the footer is made to be 36). 1.4 designates that line spacing
    # is 20% of the font size, above and below; should be just 1.2,
    # but 1.4 works better.
    page_lines = int((height - 108) / (1.4 * FONT_SIZE))
    # estimate character width using 0 as the average character
    # 100 is the left and right margins, 50 points each by default
    max_chars = int((width - 100) / fitz.Font(FONT, FONT_PATH).text_length('0', FONT_SIZE))
    line_count = 0
    page_buffer = ''

    pdf = fitz.Document()
    with open(file_name) as text_file:
        for item in text_file:
            if len(item) > max_chars:
                line_list = textwrap.wrap(item, width=max_chars, subsequent_indent='\n')
                line_list[-1] += '\n'
            else:
                line_list = [item]
            for line in line_list:
                page_buffer += line
                line_count += 1
                if line_count >= page_lines:
                    pdf.insert_page(
                        -1, text=page_buffer, fontsize=FONT_SIZE,
                        width=width, height=height, fontname=FONT, fontfile=FONT_PATH
                    )
                    line_count = 0
                    page_buffer = ''
        if page_buffer:
            pdf.insert_page(
                -1, text=page_buffer, fontsize=FONT_SIZE,
                width=width, height=height, fontname=FONT, fontfile=FONT_PATH
            )

    return pdf


def image_to_pdf(file_name):
    """
    Converts image files to pdf.

    Expands the image to fit PAGE_LAYOUT, while retaining its aspect ratio.

    Parameters
    ----------
    file_name : str or os.Pathlike
        The file path for the text document.

    Returns
    -------
    pdf : fitz.Document
        The file converted to a pymupdf Document.

    Notes
    -----
    Uses fitz.Page.insert_image rather than fitz.Page.show_pdf_page
    for most image types becaues insert_image correctly keeps transparency.
    show_pdf_page is used for svg images, since they cannot be directly
    used by fitz.Pixmap. Not

    """
    path = Path(file_name)
    stream = None
    with fitz.Document(str(path)) as doc:
        image_rect = doc[0].rect
        if 'svg' in path.suffix:
            # have to convert svg to pdf since pymupdf cannot use svg as a Pixmap
            stream = doc.convert_to_pdf()

    if EXPAND_IMAGES or FULL_PAGE_IMAGES:
        width, height = fitz.PaperSize(get_page_layout())
    else:
        width, height = image_rect.width, image_rect.height

    pdf = fitz.Document()
    page = pdf.new_page(width=width, height=height)

    if EXPAND_IMAGES:
        rect = page.rect
    elif not FULL_PAGE_IMAGES:
        rect = image_rect
    else:
        if image_rect.width > image_rect.height:
            rect_width = min(image_rect.width, page.rect.width)
            rect_height = (rect_width / image_rect.width) * image_rect.height
        else:
            rect_height = min(image_rect.height, page.rect.height)
            rect_width = (rect_height / image_rect.height) * image_rect.width

        rect = fitz.Rect(0, 0, rect_width, rect_height)
        if CENTER_IMAGES_H:
            rect += ((page.rect.width - rect_width) / 2, 0, (page.rect.width - rect_width) / 2, 0)
        if CENTER_IMAGES_V:
            rect += (
                0, (page.rect.height - rect_height) / 2, 0, (page.rect.height - rect_height) / 2
            )

    if stream is not None:
        with fitz.Document(filetype='pdf', stream=stream) as image_pdf:
            page.show_pdf_page(rect, image_pdf, 0)
    else:
        page.insert_image(rect, filename=str(path))

    return pdf


def is_image(file_suffix):
    """
    Determines if the file is a supported image file for pymupdf.

    Parameters
    ----------
    file_suffix : str
        The file extension for the document (eg. 'jpg', 'pdf'). Does not matter
        if preceeded by a period (ie. '.jpg' and 'jpg' are both recognized).

    Returns
    -------
    bool
        Returns True if the file's extension is recognized as a compatible
        image type for pymupdf.

    Notes
    -----
    Supported extensions are jpg/jpeg/jpx, png, tif/tiff, svg, gif, and bmp.

    """
    return re.search('jp.*|png|tif.*|svg|gif|bmp', file_suffix) is not None


class PDFViewer(wx.Frame):
    """
    A frame for displaying files using pymupdf.

    Parameters
    ----------
    parent : wx.Window
        The parent widget for the frame.
    pdf : str or io.BytesIO or os.Pathlike or fitz.Document
        The file or buffer stream to display.
    **kwargs
        Any additional keyword arguments for initializing wx.Frame.

    """

    def __init__(self, parent, pdf, **kwargs):
        super().__init__(parent, **kwargs)
        self.current_pg = 1
        self.zoom_level = -1  # fit page width
        self.matrix = fitz.Matrix(1, 1)
        self.pdf = self.load_pdf(pdf)
        self._total_pages = len(self.pdf)
        self._last_v_scroll = -1  # used to track vertical scrolling
        self._redraw = False

        if LOGO is not None:
            self.SetIcon(wx.Icon(wx.Image(io.BytesIO(base64.b64decode(LOGO))).ConvertToBitmap()))

        self.main_panel = wx.Panel(self)
        sizer_1 = wx.BoxSizer(wx.VERTICAL)

        self.tool_panel = wx.Panel(self.main_panel)
        sizer_1.Add(self.tool_panel, 0, wx.ALL | wx.EXPAND, 5)

        sizer_2 = wx.BoxSizer(wx.HORIZONTAL)

        self.back_btn = wx.Button(self.tool_panel, label='Back', style=wx.BU_EXACTFIT)
        sizer_2.Add(self.back_btn, 0, wx.LEFT, 5)

        self.next_btn = wx.Button(self.tool_panel, label='Next', style=wx.BU_EXACTFIT)
        sizer_2.Add(self.next_btn, 0, wx.RIGHT, 5)

        # filler
        sizer_2.Add((5, 5), 1, 0, 0)

        label_2 = wx.StaticText(self.tool_panel, label='Page: ')
        sizer_2.Add(label_2, 0, wx.ALIGN_CENTER_VERTICAL | wx.LEFT, 20)

        # creates a temporary wx.StaticText to determine the minimum size
        # for the text control so that it isn't too large.
        self.pg_input = wx.TextCtrl(self.tool_panel, value='1', style=wx.TE_PROCESS_ENTER)
        temp_label = wx.StaticText(self.tool_panel, label='11111111')
        self.pg_input.SetMinSize((temp_label.GetSize()[0], -1))
        temp_label.Destroy()
        temp_label = None
        sizer_2.Add(self.pg_input, 0, wx.ALL, 0)

        total_pages = wx.StaticText(self.tool_panel, label=f' /{self._total_pages}')
        sizer_2.Add(total_pages, 0, wx.ALIGN_CENTER_VERTICAL | wx.RIGHT, 5)

        # filler
        sizer_2.Add((5, 5), 1, 0, 0)

        label_4 = wx.StaticText(self.tool_panel, label='Zoom: ')
        sizer_2.Add(label_4, 0, wx.ALIGN_CENTER_VERTICAL | wx.LEFT, 20)

        self.zoom_choice = wx.Choice(self.tool_panel)
        zoom_levels = (
            ('Fit width', -1), ('Fit page', 0), ('25%', 0.25), ('50%', 0.5),
            ('75%', 0.75), ('100%', 1.0), ('125%', 1.25), ('150%', 1.5),
            ('200%', 2.0), ('400%', 4.0)
        )
        for level in zoom_levels:
            # add both a display text and client data for doing the zooming
            self.zoom_choice.Append(*level)
        self.zoom_choice.SetSelection(0)  # default to fit width
        sizer_2.Add(self.zoom_choice, 0, 0, 0)

        self.zoom_out_btn = wx.Button(
            self.tool_panel, label='Zoom Out', style=wx.BU_EXACTFIT
        )
        sizer_2.Add(self.zoom_out_btn, 0, wx.LEFT, 10)

        self.zoom_in_btn = wx.Button(
            self.tool_panel, label='Zoom In', style=wx.BU_EXACTFIT
        )
        sizer_2.Add(self.zoom_in_btn, 0, wx.RIGHT, 5)

        self.tool_panel.SetSizer(sizer_2)

        self.display_panel = wx.ScrolledWindow(self.main_panel, style=wx.BORDER_SUNKEN)
        self.display_panel.SetScrollRate(20, 20)
        # use double buffer to prevent flickering when
        # moving between pages or scrolling on a page
        self.display_panel.SetDoubleBuffered(True)
        # initially disable horizontal scrolling
        self.display_panel.ShowScrollbars(wx.SHOW_SB_NEVER, wx.SHOW_SB_DEFAULT)
        self.display_panel.EnableScrolling(False, True)
        sizer_1.Add(self.display_panel, 1, wx.ALL | wx.EXPAND, 0)

        sizer_3 = wx.BoxSizer(wx.HORIZONTAL)
        self.pdf_bitmap = wx.StaticBitmap(self.display_panel)
        sizer_3.Add(self.pdf_bitmap, 1)
        self.render_page(1)
        self.display_panel.SetSizer(sizer_3)
        self.main_panel.SetSizer(sizer_1)

        self.Layout()
        self.tool_panel.Fit()
        self.SetSize((self.tool_panel.GetSize()[0] + 20, wx.GetDisplaySize()[1] - 100))

        self.back_btn.Bind(wx.EVT_BUTTON, self.on_back)
        self.next_btn.Bind(wx.EVT_BUTTON, self.on_next)
        self.zoom_choice.Bind(wx.EVT_CHOICE, self.on_zoom)
        self.zoom_in_btn.Bind(wx.EVT_BUTTON, self.on_zoom_in)
        self.zoom_out_btn.Bind(wx.EVT_BUTTON, self.on_zoom_out)
        self.display_panel.Bind(wx.EVT_MOUSEWHEEL, self.on_scroll)
        self.display_panel.Bind(wx.EVT_CHAR, self.on_key)
        self.display_panel.Bind(wx.EVT_LEFT_DOWN, self.set_focus)
        self.pdf_bitmap.Bind(wx.EVT_LEFT_DOWN, self.set_focus)
        self.pg_input.Bind(wx.EVT_TEXT_ENTER, self.go_to_page)
        self.pg_input.Bind(wx.EVT_KILL_FOCUS, self.go_to_page)
        self.Bind(wx.EVT_CLOSE, self.on_close)
        self.Bind(wx.EVT_SIZE, self.on_resize)
        self.Bind(wx.EVT_IDLE, self.on_idle)

    def set_focus(self, event=None):
        """Focuses on the display on left-clicks so that arrow keys can scroll."""
        self.display_panel.SetFocus()
        if event is not None:
            event.Skip()

    def on_key(self, event):
        """Allows using arrow keys to also scroll and/or move pages."""
        mouse_wheel_events = {
            wx.WXK_UP: {'axis': wx.MOUSE_WHEEL_VERTICAL, 'rotation': 1},
            wx.WXK_DOWN: {'axis': wx.MOUSE_WHEEL_VERTICAL, 'rotation': -1},
            wx.WXK_LEFT: {'axis': wx.MOUSE_WHEEL_HORIZONTAL, 'rotation': -1},
            wx.WXK_RIGHT: {'axis': wx.MOUSE_WHEEL_HORIZONTAL, 'rotation': 1},
        }
        if event.GetKeyCode() not in mouse_wheel_events:
            event.Skip()
        else:
            mouse_event = wx.MouseEvent()
            key = event.GetKeyCode()
            mouse_event.SetWheelAxis(mouse_wheel_events[key]['axis'])
            mouse_event.SetWheelRotation(mouse_wheel_events[key]['rotation'])

            page = self.current_pg
            self.on_scroll(mouse_event)
            if self.current_pg == page:  # scroll only if pages were not changed
                event.Skip()

    def on_idle(self, event):
        """Redraws the image once sizing is completed."""
        if self._redraw:
            self.render_page(self.current_pg)
            self._redraw = False
        event.Skip()

    def on_resize(self, event):
        """Queues up a redraw of the pdf if zoom level is set to fit page or width."""
        if not self._redraw and self.zoom_level <= 0:
            self._redraw = True
        event.Skip()

    def load_pdf(self, pdf):
        """
        Creates a pymupdf Document for viewing.

        Parameters
        ----------
        pdf : str or io.BytesIO or os.Pathlike or fitz.Document
            The file or buffer stream to display.

        Returns
        -------
        fitz.Document
            The fitz Document.

        Notes
        -----
        If the input pdf is already a fitz.Document, then a new document is created
        from that document so that this frame has sole ownership of the pdf.

        """
        if isinstance(pdf, fitz.Document):
            # fitz.Document.write was changes to .tobytes in pymupdf v1.18.7
            if hasattr(pdf, 'tobytes'):
                stream = io.BytesIO(pdf.tobytes())
            else:
                stream = io.BytesIO(pdf.write())
            file_name = 'pdf'
        elif isinstance(pdf, (str, os.PathLike)):
            file_name = str(Path(pdf))
            stream = None
        else:
            # assume it is a file stream
            file_name = 'pdf'
            stream = pdf

        return fitz.Document(file_name, stream)

    @functools.lru_cache(maxsize=500)
    def get_displaylist(self, page):
        """
        Gets the DisplayList for the given page.

        Results are cached so that repeated lookups are faster (The DisplayLists
        do not require much memory to cache).

        Parameters
        ----------
        page : int
            The page to read from the pdf. 1-based.

        Returns
        -------
        fitz.DisplayList
            The DisplayList for the page.

        """
        return self.pdf[page - 1].get_displaylist()

    def render_page(self, page, reset_scroll=False):
        """
        Displays the selected page onto the frame.

        Parameters
        ----------
        page : int
            The page to render. Is 1-indexed.
        reset_scroll : bool
            If False (default), will place the scrollbar at the bottom of the page
            when going back pages. Otherwise, will place the scrollbar at the top.

        Notes
        -----
        If the zoom level is set to fit the page, then the zoom matrix
        is recalculated for each page.

        Use self.Freeze to not update the frame while changing pages so
        that the transition is much smoother, and then call self.Thaw.

        """
        if self.zoom_level <= 0:
            area = self.get_displaylist(page).rect
            height = area.height
            width = area.width
            display_area = self.display_panel.GetSize()
            if self.zoom_level == 0:
                zoom = min(display_area[0] / width, display_area[1] / height)
            else:
                # give a little space for a possible scrollbar
                zoom = (display_area[0] - self.FromDIP(25)) / width
            self.matrix = fitz.Matrix(zoom, zoom)

        self.Freeze()
        pixmap = self.get_displaylist(page).getPixmap(matrix=self.matrix, alpha=False)
        self.pdf_bitmap.SetBitmap(wx.Bitmap.FromBuffer(pixmap.w, pixmap.h, pixmap.samples))
        pixmap = None

        self.main_panel.Layout()  # updates the display_panel's scrollbars
        # update scrollbar position
        if self.display_panel.CanScroll(wx.VERTICAL) and not reset_scroll:
            v_scroll = self.display_panel.GetScrollRange(wx.VERTICAL)
        else:
            v_scroll = 0

        if self.current_pg == self._total_pages and page == 1:
            self.display_panel.Scroll(0, 0)
        elif self.current_pg == 1 and page == self._total_pages:
            self.display_panel.Scroll(0, v_scroll)
        elif page > self.current_pg:
            self.display_panel.Scroll(0, 0)
        elif page < self.current_pg:
            self.display_panel.Scroll(0, v_scroll)

        self.current_pg = page
        self._last_v_scroll = -1
        self.display_panel.SetFocus()

        self.Thaw()

    def go_to_page(self, event):
        """Goes to the page specified by the user, if it is valid."""
        try:
            page = int(self.pg_input.GetValue())
        except ValueError:
            page = self.current_pg
            self.pg_input.SetValue(str(self.current_pg))

        if page != self.current_pg:
            self._go_to_page(page)
            self.display_panel.Scroll(0, 0)
        event.Skip()

    def on_back(self, event=None):
        """Goes to the previous page and resets the scrollbar."""
        self._go_to_page(self.current_pg - 1, True)

    def on_next(self, event=None):
        """Goes to the next page and resets the scrollbar."""
        self._go_to_page(self.current_pg + 1, True)

    def _go_to_page(self, page, reset_scroll=False):
        """
        Goes to the selected page.

        Parameters
        ----------
        page : int
            The page to render; 1-based. If the specified page is < 1, then the last
            page of the document is rendered. Likewise, if the specified page is greater
            than the total number of pages, then the first page is shown. This way, can
            scroll from the first page back to the last page.
        reset_scroll : bool


        """
        if page > self._total_pages:
            new_page = 1
        elif page < 1:
            new_page = self._total_pages
        else:
            new_page = page

        self.pg_input.SetValue(str(new_page))
        self.render_page(new_page, reset_scroll)

    def on_zoom(self, event=None):
        """Gets the zoom level specified by the user."""
        zoom_level = self.zoom_choice.GetClientData(self.zoom_choice.GetSelection())
        self._zoom(zoom_level)
        event.Skip()

    def on_zoom_out(self, event):
        """Increases the zoom level, if possible."""
        new_zoom = None
        zoom_index = self.zoom_choice.GetSelection()
        if self.zoom_choice.GetClientData(zoom_index) <= 0:
            new_zoom = 1  # reset to 100% zoom
            new_index = 5
        elif self.zoom_choice.GetClientData(zoom_index - 1) > 0:
            new_zoom = self.zoom_choice.GetClientData(zoom_index - 1)
            new_index = zoom_index - 1
        if new_zoom:
            self.zoom_choice.SetSelection(new_index)
            self._zoom(new_zoom)

        event.Skip()

    def on_zoom_in(self, event):
        """Decreases the zoom level, if possible."""
        new_zoom = None
        zoom_index = self.zoom_choice.GetSelection()
        if self.zoom_choice.GetClientData(zoom_index) <= 0:
            new_zoom = 1  # reset to 100% zoom
            new_index = 5
        elif zoom_index + 1 < self.zoom_choice.GetCount():
            new_zoom = self.zoom_choice.GetClientData(zoom_index + 1)
            new_index = zoom_index + 1
        if new_zoom:
            self.zoom_choice.SetSelection(new_index)
            self._zoom(new_zoom)
        event.Skip()

    def _zoom(self, zoom_level):
        """
        Updates the zoom level and renders the page with the new zoom.

        If the zoom level is 0 (fit to page), then scrolling is disabled and the
        scrollbars are made to not appear. If the zoom level if -1 (fit to width),
        then only horizontal scrollbars are disabled and hidden. Otherwise, the
        scrollbars are enabled and will appear if needed.

        """
        if zoom_level == self.zoom_level:
            return

        if zoom_level <= 0:
            self.display_panel.ShowScrollbars(
                wx.SHOW_SB_NEVER, wx.SHOW_SB_DEFAULT if zoom_level == -1 else wx.SHOW_SB_NEVER
            )
            self.display_panel.EnableScrolling(False, zoom_level == -1)
        else:
            self.display_panel.ShowScrollbars(wx.SHOW_SB_DEFAULT, wx.SHOW_SB_DEFAULT)
            self.display_panel.EnableScrolling(True, True)
        self.zoom_level = zoom_level
        if self.zoom_level > 0:
            self.matrix = fitz.Matrix(self.zoom_level, self.zoom_level)
        self.render_page(self.current_pg)

    def on_scroll(self, event):
        """
        Handles scrolling up or down to either scroll or move pages.

        Horizontal scrolls are ignored, and will not move between pages.

        """
        if event.GetWheelAxis() == wx.MOUSE_WHEEL_HORIZONTAL:
            event.Skip()
            return

        if not (self.display_panel.HasScrollbar(wx.VERTICAL)
                and self.display_panel.GetScrollPos(wx.VERTICAL) != self._last_v_scroll):
            # scroll up gives positive rotation
            if event.GetWheelRotation() > 0:
                self._go_to_page(self.current_pg - 1)
            else:
                self._go_to_page(self.current_pg + 1)
        else:
            event.Skip()  # only skip if not changing pages so it doesn't scroll
            self._last_v_scroll = self.display_panel.GetScrollPos(wx.VERTICAL)

    def on_close(self, event):
        """
        Cleans up everything associated with the window before closing.

        Closes the fitz.Document and clears the DisplayList cache.

        """
        try:
            self.pdf.close()
        except Exception:
            pass  # the fitz.Document was already closed
        self.pdf = None
        self.get_displaylist.cache_clear()
        event.Skip()


class PDFApp(wx.App):
    """The app for launching the PDFMerger frame."""

    def OnInit(self):
        """Handles the initialization of the PDFMerger frame."""
        self.frame = PDFMerger(None, title='PDF Merger')
        self.SetTopWindow(self.frame)
        self.frame.Show()
        return True


if __name__ == "__main__":

    # Set dpi awareness on Windows operating systems.
    if os.name == 'nt':
        try:
            import ctypes
            ctypes.oledll.shcore.SetProcessDpiAwareness(1)
        except (AttributeError, ImportError, OSError, PermissionError):
            pass

    app = PDFApp()
    app.MainLoop()
