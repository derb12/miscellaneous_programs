# -*- coding: utf-8 -*-
"""A GUI for selecting PDF files to merge into one file and save.

Also allows selecting the first page, last page, and rotation of each individual file
and allows previewing the merged file before saving.

@author: Donald Erb
Created on 2021-02-04 14:56:34

Requirements
------------
fitz>=1.18.4
    fitz is pymupdf; new versions have this_case rather than CamelCase or mixedCase.
wx>=4.0
    wx is wxpython; v4.0 is the first to work with python v3.

License:
--------
This program is licensed under the GNU GPL V3+.

The set_dpi_awareness function was copied from mcetl. The mcetl license is included below:

BSD 3-Clause License

Copyright (c) 2020-2021, Donald Erb
All rights reserved.

Redistribution and use in source and binary forms, with or without
modification, are permitted provided that the following conditions are met:

1. Redistributions of source code must retain the above copyright notice, this
   list of conditions and the following disclaimer.

2. Redistributions in binary form must reproduce the above copyright notice,
   this list of conditions and the following disclaimer in the documentation
   and/or other materials provided with the distribution.

3. Neither the name of the copyright holder nor the names of its
   contributors may be used to endorse or promote products derived from
   this software without specific prior written permission.

THIS SOFTWARE IS PROVIDED BY THE COPYRIGHT HOLDERS AND CONTRIBUTORS "AS IS"
AND ANY EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT LIMITED TO, THE
IMPLIED WARRANTIES OF MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE ARE
DISCLAIMED. IN NO EVENT SHALL THE COPYRIGHT HOLDER OR CONTRIBUTORS BE LIABLE
FOR ANY DIRECT, INDIRECT, INCIDENTAL, SPECIAL, EXEMPLARY, OR CONSEQUENTIAL
DAMAGES (INCLUDING, BUT NOT LIMITED TO, PROCUREMENT OF SUBSTITUTE GOODS OR
SERVICES; LOSS OF USE, DATA, OR PROFITS; OR BUSINESS INTERRUPTION) HOWEVER
CAUSED AND ON ANY THEORY OF LIABILITY, WHETHER IN CONTRACT, STRICT LIABILITY,
OR TORT (INCLUDING NEGLIGENCE OR OTHERWISE) ARISING IN ANY WAY OUT OF THE USE
OF THIS SOFTWARE, EVEN IF ADVISED OF THE POSSIBILITY OF SUCH DAMAGE.

"""

import io
import os
from pathlib import Path
import traceback

import fitz
import wx
import wx.grid
from wx.lib.pdfviewer import pdfViewer, pdfButtonPanel

try:
    import ctypes
except ImportError:
    ctypes = None


fitz.TOOLS.mupdf_display_errors(False)


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

    def add_row(self, file_path):
        """
        Appends a row to the grid with information after opening the given file.

        Parameters
        ----------
        file_path : str or os.Pathlike
            The path of the pdf file to open and read.

        """
        try:
            pdf = fitz.Document(file_path)
            if pdf.needsPass:  # password protected
                raise ValueError('File is encrypted and cannot be processed')
            pages = len(pdf)
        except Exception as exc:
            with wx.MessageDialog(
                self,
                (f'Problem opening {file_path}\n\nError:\n{repr(exc)}'),
            ) as dlg:
                dlg.ShowModal()
            return
        finally:
            try:
                pdf.close()
            except Exception:
                pass

        row = self.GetNumberRows()
        self.AppendRows(1)
        self.create_row(row, file_path, pages)

    def create_row(self, row, file_path, total_pages, first_page='1', last_page=None, rotation=None):
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
        self.SetSize((990, 520))
        self.preview = None

        self.panel = wx.Panel(self)
        sizer = wx.BoxSizer(wx.VERTICAL)

        sizer_1 = wx.BoxSizer(wx.HORIZONTAL)
        sizer.Add(sizer_1, 0, wx.EXPAND, 0)

        grid_sizer_1 = wx.GridSizer(1, 2, 0, 20)
        sizer_1.Add(grid_sizer_1, 0, wx.ALIGN_CENTER_VERTICAL | wx.ALL, 10)

        self.add_btn = wx.Button(self.panel, wx.ID_ANY, 'Add Files', style=wx.BU_EXACTFIT)
        grid_sizer_1.Add(self.add_btn, 0, wx.EXPAND, 0)

        self.remove_btn = wx.Button(self.panel, wx.ID_ANY, 'Remove Files', style=wx.BU_EXACTFIT)
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

        self.preview_btn = wx.Button(self.panel, wx.ID_ANY, 'Preview')
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

    def on_add(self, event):
        """Launches file dialog to select pdf files and adds them to the grid."""
        paths = []
        with wx.FileDialog(
            self, 'Select files',
            wildcard='PDF Files (*.pdf)|*.pdf',
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
            output_pdf.save(output_path, garbage=4, deflate=1)
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
            self.Close()
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
    output_file = fitz.Document()  # same as fitz.open, but usage is more clear
    total_toc = []  # collects the bookmarks from all of the files
    for file_path, first_pg, last_pg, rotation in grid_data:
        path = Path(file_path)
        if path.suffix.lower() != '.pdf':
            print(f'\nThe following file cannot be processed:\n\n    {str(path)}')
            continue

        with fitz.open(str(path)) as temp_file:
            pages = len(temp_file)

            # ensures pages are within the document
            first_pg = min(max(0, first_pg), pages - 1)
            if last_pg != -1:
                last_pg = min(max(0, last_pg), pages - 1)
            else:
                last_pg = pages - 1

            # only add unencrypted files
            if not temp_file.needsPass:
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

                # skip the bookmark if it's "goto" and not within the page range
                if (link[2] - 1) not in pg_range and lnk_type == fitz.LINK_GOTO:
                    continue
                elif lnk_type == fitz.LINK_GOTO:
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


class PDFApp(wx.App):
    """The app for launching the PDFMerger frame."""

    def OnInit(self):
        """Handles the initialization of the PDFMerger frame."""
        self.frame = PDFMerger(None, title='PDF Merger')
        self.SetTopWindow(self.frame)
        self.frame.Show()
        return True


def set_dpi_awareness(awareness_level=1):
    """
    Sets DPI awareness for Windows operating system so that GUIs are not blurry.

    Fixes blurry GUIs due to weird dpi scaling in Windows os. Other
    operating systems are ignored.

    Parameters
    ----------
    awareness_level : {1, 0, 2}
        The dpi awareness level to set. 0 turns off dpi awareness, 1 sets dpi
        awareness to scale with the system dpi and automatically changes when
        the system dpi changes, and 2 sets dpi awareness per monitor and does
        not change when system dpi changes. Default is 1.

    Raises
    ------
    ValueError
        Raised if awareness_level is not 0, 1, or 2.

    Notes
    -----
    Will only work on Windows 8.1 or Windows 10. Not sure if earlier versions
    of Windows have this issue anyway.

    Copied from mcetl.

    """
    # 'nt' designates Windows operating system
    if os.name == 'nt' and ctypes is not None:
        if awareness_level not in (0, 1, 2):
            raise ValueError('Awareness level must be either 0, 1, or 2.')
        try:
            ctypes.oledll.shcore.SetProcessDpiAwareness(awareness_level)
        except (AttributeError, OSError, PermissionError):
            # AttributeError is raised if the dll loader was not created, OSError
            # is raised if setting the dpi awareness errors, and PermissionError is
            # raised if the dpi awareness was already set, since it can only be set
            # once per thread. All are ignored.
            pass


class ButtonPanel(pdfButtonPanel):
    """
    A custom button panel that overrides some of the buttons from pdfButtonPanel.

    Ensures saving and printing is not done through the viewer, and only allows
    zooming up to 200% (can still manually enter a higher zoom number though, but
    that requires deliberate action to do).

    Parameters
    ----------
    parent : wx.Window
        The parent widget for the button panel.

    """

    def __init__(self, parent):
        super().__init__(parent, wx.ID_ANY, wx.DefaultPosition, wx.DefaultSize, 0)

        # remove zooms > 200% since high zooms greatly increases memory usage
        # and lags the viewer.
        delete_indices = []
        for i, selection in enumerate(self.zoom.GetStrings()):
            try:
                if int(selection[:-1]) > 200:
                    delete_indices.append(i)
            except ValueError:
                pass
        for i in reversed(delete_indices):
            self.zoom.Delete(i)
        self.DoLayout()

    def OnSave(self, event):
        """Overrides super class's method to disable saving."""
        with wx.MessageDialog(
            self.GetParent(),
            'Saving is not allowed here.', 'Saving Disabled'
        ) as dlg:
            dlg.ShowModal()

    def OnPrint(self, event):
        """Overrides super class's method to disable printing."""
        with wx.MessageDialog(
            self.GetParent(),
            'Printing is not allowed here.', 'Printing Disabled'
        ) as dlg:
            dlg.ShowModal()

    def OnZoomIn(self, event):
        """Overrides super class's method to ensure max zoom using the button is 200%."""
        if self.percentzoom > 100:
            self.percentzoom = 100
        super().OnZoomIn(event)


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

        self.panel = wx.Panel(self)
        self.sizer = wx.BoxSizer(wx.VERTICAL)
        self.panel.SetSizer(self.sizer)
        self.buttons = ButtonPanel(self.panel)

        self.viewer = pdfViewer(
            self.panel, wx.ID_ANY, wx.DefaultPosition, wx.DefaultSize, wx.BORDER_SUNKEN
        )
        self.viewer.SetDoubleBuffered(True)  # reduces flicker when scrolling

        # let the viewer and buttons recognize each other
        self.buttons.viewer = self.viewer
        self.viewer.buttonpanel = self.buttons

        self.sizer.Add(self.buttons, 0, wx.EXPAND)
        self.sizer.Add(self.viewer, 1, wx.EXPAND)

        self.load_pdf(pdf)
        self.sizer.Fit(self.buttons)
        self.SetSize((self.buttons.GetSize()[0], wx.GetDisplaySize()[1] - 100))

        self.Bind(wx.EVT_CLOSE, self.on_close)

    def load_pdf(self, pdf):
        """
        Ensures that the pdf is a str or buffer so that pdfViewer can use it.

        Parameters
        ----------
        pdf : str or io.BytesIO or os.Pathlike or fitz.Document
            The file or buffer stream to display.

        """
        if isinstance(pdf, (str, io.BytesIO)):
            file_stream = pdf
        elif isinstance(pdf, os.PathLike):
            file_stream = str(Path(pdf))
        elif isinstance(pdf, fitz.Document):
            # fitz.Document.write was changes to .tobytes in pymupdf v1.18.7
            if hasattr(pdf, 'tobytes'):
                file_stream = io.BytesIO(pdf.tobytes())
            else:
                file_stream = io.BytesIO(pdf.write())
        else:
            raise NotImplementedError(f'Using {type(pdf)} is not implemented.')

        self.viewer.LoadFile(file_stream)

    def on_close(self, event):
        """Ensures that the fitz.Document is closed."""
        try:
            self.viewer.pdfdoc.pdfdoc.close()
        except ValueError:
            pass  # the fitz.Document was already closed
        event.Skip()


if __name__ == "__main__":
    set_dpi_awareness()
    app = PDFApp()
    app.MainLoop()
