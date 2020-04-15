# encoding: utf-8

"""
Paragraph-related proxy types.
"""

from __future__ import (
    absolute_import, division, print_function, unicode_literals
)

from ..enum.style import WD_STYLE_TYPE
from .parfmt import ParagraphFormat
from .run import Run, Del
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

    def add_del_at_start(self):
        """
        Insert a del element at the start of this paragraph
        """
        d = self._p.add_d()
        d_elem = Del(d, self)

        return d_elem

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

        return None, None


    def find_text_end_run(self, text, start_run=None, start_index=None):
        """
        Return the run and index within the run that indicates where
        the text provided, ends.
        """

        if start_run:
            if len(text) <= len(start_run.text) - start_index:
                return start_run, start_run.text.index(text) + len(text)
            else:
                text = text[len(start_run.text)-start_index:]
                current_run = start_run
                while (current_run.next()):
                    current_run = current_run.next()
                    if current_run.text and text in current_run.text:
                        return current_run, len(text) - 1
                    elif current_run.text:
                        text = text[len(current_run.text):]

        else:
            for run in self.runs:
                end_index = run.text_end_index(text)
                if end_index is not None:
                    return run, end_index

        return None, None

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

    def create_track_change_elements(self, start_run, end_run, create_ins_element=True):
        """
        Create insertion and deletion elements within the paragraph, depending
        on whether the start and end runs containing text to be replaced, are
        present within this paragraph.
        """

        del_element = ins_element = None

        if start_run and not end_run:
            # Text starts but doesn't end in this paragraph
            del_element = start_run.add_track_delete_after()
        elif start_run and end_run:
            # Text starts and ends in this paragraph
            del_element = start_run.add_track_delete_after()
            if create_ins_element:
                ins_element = start_run.add_track_insert_after()
        elif end_run:
            # Text ends but does not start in this paragraph
            if create_ins_element:
                ins_element = end_run.add_track_insert_before()
            del_element = end_run.add_track_delete_before()
        else:
            # Text doesn't start or end in this paragraph
            del_element = self.add_del_at_start()

        return ins_element, del_element


    def replace_text_track_change(self, original_text, replacement_text):
        """
        Replace text with some replacement text. Determine whether (and where)
        the text starts within the paragraph, and add tracked revisions to
        indicate deletions and insertions.
        """

        start_run, start_index = self.find_text_start_run(original_text)
        end_run, end_index = self.find_text_end_run(original_text, start_run, start_index)

        ins_element, del_element = self.create_track_change_elements(\
            start_run, end_run, create_ins_element=replacement_text != None)
        replace_start = True if not start_run else False

        for run in self.runs:
            if start_run and end_run and run.text == start_run.text and start_run.text == end_run.text:
                del_element.add_deltext(run, run.text[start_index:end_index])
                split_run = del_element.add_run_after(original_run=start_run)
                split_run.text = run.text[end_index:]
                run.text = run.text[:start_index]
            elif start_run and run.text == start_run.text:
                replace_start = True
                del_element.add_deltext(run, run.text[start_index:])
                run.text = run.text[:start_index]
            elif replace_start and end_run and run.text == end_run.text:
                del_element.add_deltext(run, run.text[:end_index+1])
                split_run = del_element.add_run_after(original_run=end_run)
                split_run.text = run.text[end_index+1:]
                run.text = ""
                break
            elif replace_start:
                del_element.add_deltext(run, run.text)
                self.delete_run(run)

        if ins_element:
            ins_element.add_run(replacement_text, original_run=end_run)

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

            # Remove the last word from the text
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

            # Remove the last character of the text
            text = " ".join(text.split(" ")[1:])

        return False

    def _insert_paragraph_before(self):
        """
        Return a newly created paragraph, inserted directly before this
        paragraph.
        """
        p = self._p.add_p_before()
        return Paragraph(p, self._parent)
