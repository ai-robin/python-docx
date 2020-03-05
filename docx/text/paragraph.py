# encoding: utf-8

"""
Paragraph-related proxy types.
"""

from __future__ import (
    absolute_import, division, print_function, unicode_literals
)

from ..enum.style import WD_STYLE_TYPE
from .parfmt import ParagraphFormat
from .run import Run
from ..shared import Parented


class Paragraph(Parented):
    """
    Proxy object wrapping ``<w:p>`` element.
    """
    def __init__(self, p, parent):
        super(Paragraph, self).__init__(parent)
        self._p = self._element = p

    def add_run(self, text=None, style=None):
        """
        Append a run to this paragraph containing *text* and having character
        style identified by style ID *style*. *text* can contain tab
        (``\\t``) characters, which are converted to the appropriate XML form
        for a tab. *text* can also include newline (``\\n``) or carriage
        return (``\\r``) characters, each of which is converted to a line
        break.
        """
        r = self._p.add_r()
        run = Run(r, self)
        if text:
            run.text = text
        if style:
            run.style = style
        return run

    @property
    def alignment(self):
        """
        A member of the :ref:`WdParagraphAlignment` enumeration specifying
        the justification setting for this paragraph. A value of |None|
        indicates the paragraph has no directly-applied alignment value and
        will inherit its alignment value from its style hierarchy. Assigning
        |None| to this property removes any directly-applied alignment value.
        """
        return self._p.alignment

    @alignment.setter
    def alignment(self, value):
        self._p.alignment = value

    def clear(self):
        """
        Return this same paragraph after removing all its content.
        Paragraph-level formatting, such as style, is preserved.
        """
        self._p.clear_content()
        return self

    def delete_run(self, run):
        self._p.remove(run._r)

    def find_text_start_run(self, text):
        """
        Return the run and index within the run that indicates where
        the text provided, starts.
        """

        for run in self.runs:
            start_index = run.text_start_index(text)
            if start_index is not None:
                return run, start_index


    def find_text_end_run(self, text):
        """
        Return the run and index within the run that indicates where
        the text provided, ends.
        """

        for run in self.runs:
            end_index = run.text_end_index(text)
            if end_index is not None:
                return run, end_index


    def insert_paragraph_before(self, text=None, style=None):
        """
        Return a newly created paragraph, inserted directly before this
        paragraph. If *text* is supplied, the new paragraph contains that
        text in a single run. If *style* is provided, that style is assigned
        to the new paragraph.
        """
        paragraph = self._insert_paragraph_before()
        if text:
            paragraph.add_run(text)
        if style is not None:
            paragraph.style = style
        return paragraph

    @property
    def paragraph_format(self):
        """
        The |ParagraphFormat| object providing access to the formatting
        properties for this paragraph, such as line spacing and indentation.
        """
        return ParagraphFormat(self._element)

    def replace_text_track_change(self, original_text, replacement_text):
        """

        """

        start_run = self.find_text_start_run(original_text)
        end_run = self.find_text_end_run(original_text)
        del_element = start_run[0].add_track_delete_after()
        ins_element = start_run[0].add_track_insert_after()
        found_start = False
        for run in self.runs:
            if run.text == start_run[0].text:
                found_start = True
                del_element.add_deltext(run, run.text[start_run[1]:])
                run.text = run.text[:start_run[1]]
            elif end_run and run.text == end_run[0].text:
                del_element.add_deltext(run, run.text[:end_run[1]+1])
                run.text = run.text[end_run[1]+1:]
                break
            elif found_start:
                del_element.add_deltext(run, run.text)
                self.delete_run(run)

        ins_element.add_run(replacement_text)

    @property
    def runs(self):
        """
        Sequence of |Run| instances corresponding to the <w:r> elements in
        this paragraph.
        """
        return [Run(r, self) for r in self._p.r_lst]

    @property
    def style(self):
        """
        Read/Write. |_ParagraphStyle| object representing the style assigned
        to this paragraph. If no explicit style is assigned to this
        paragraph, its value is the default paragraph style for the document.
        A paragraph style name can be assigned in lieu of a paragraph style
        object. Assigning |None| removes any applied style, making its
        effective value the default paragraph style for the document.
        """
        style_id = self._p.style
        return self.part.get_style(style_id, WD_STYLE_TYPE.PARAGRAPH)

    @style.setter
    def style(self, style_or_name):
        style_id = self.part.get_style_id(
            style_or_name, WD_STYLE_TYPE.PARAGRAPH
        )
        self._p.style = style_id

    @property
    def text(self):
        """
        String formed by concatenating the text of each run in the paragraph.
        Tabs and line breaks in the XML are mapped to ``\\t`` and ``\\n``
        characters respectively.

        Assigning text to this property causes all existing paragraph content
        to be replaced with a single run containing the assigned text.
        A ``\\t`` character in the text is mapped to a ``<w:tab/>`` element
        and each ``\\n`` or ``\\r`` character is mapped to a line break.
        Paragraph-level formatting, such as style, is preserved. All
        run-level formatting, such as bold or italic, is removed.
        """
        text = ''
        for run in self.runs:
            text += run.text
        return text

    @text.setter
    def text(self, text):
        self.clear()
        self.add_run(text)

    def text_starts_in_paragraph(self, text):
        """
        Return a boolean representing whether the text provided
        starts within the paragraph.
        """

        if text in self.text:
            return True

        while text:
            if self.text.endswith(text):
                return True

            text = " ".join(text.split(" ")[:len(text.split(" "))-1])

        return False

    def text_ends_in_paragraph(self, text):
        """
        Return a boolean representing whether any portion of the text provided
        ends within the paragraph.
        """

        if text in self.text:
            return True

        while text:
            if self.text.startswith(text):
                return True

            text = " ".join(text.split(" ")[1:])

        return False

    def _insert_paragraph_before(self):
        """
        Return a newly created paragraph, inserted directly before this
        paragraph.
        """
        p = self._p.add_p_before()
        return Paragraph(p, self._parent)
