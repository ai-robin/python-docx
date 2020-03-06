# encoding: utf-8

"""
Custom element classes related to paragraphs (CT_P).
"""

from ..ns import qn
from ..xmlchemy import BaseOxmlElement, OxmlElement, ZeroOrMore, ZeroOrOne


class CT_Ins(BaseOxmlElement):
    """
    ``<w:p>`` element, containing runs marked for insertion into a paragraph.
    """

    r = ZeroOrMore('w:r')

    def add_del_after(self):
        """
        Return a new ``<w:del>`` element inserted directly after this one.
        """
        new_del = OxmlElement('w:del')
        self.addnext(new_del)
        return new_del

    @property
    def text(self):
        """
        Return an Empty string, as text will be interpreted based on
        the child runs.
        """

        return ''


class CT_Del(BaseOxmlElement):
    """
    ``<w:p>`` element, containing runs marked for deletion from a paragraph.
    """

    r = ZeroOrMore('w:r')

    def add_del_after(self):
        """
        Return a new ``<w:del>`` element inserted directly after this element.
        """
        new_del = OxmlElement('w:del')
        self.addnext(new_del)
        return new_del

    def add_run_after(self):
        """
        Return a new ``<w:r>`` element inserted directly after this element
        """
        new_run = OxmlElement('w:r')
        self.addnext(new_run)
        return new_run

    @property
    def text(self):
        """
        Return an Empty string, as text should not be considered present
        within the current version of the paragraph.
        """

        return ''


class CT_P(BaseOxmlElement):
    """
    ``<w:p>`` element, containing the properties and text for a paragraph.
    """
    pPr = ZeroOrOne('w:pPr')
    r = ZeroOrMore('w:r')
    i = ZeroOrMore('w:ins')
    d = ZeroOrMore('w:del')

    def _insert_pPr(self, pPr):
        self.insert(0, pPr)
        return pPr

    def add_p_before(self):
        """
        Return a new ``<w:p>`` element inserted directly prior to this one.
        """
        new_p = OxmlElement('w:p')
        self.addprevious(new_p)
        return new_p

    @property
    def alignment(self):
        """
        The value of the ``<w:jc>`` grandchild element or |None| if not
        present.
        """
        pPr = self.pPr
        if pPr is None:
            return None
        return pPr.jc_val

    @alignment.setter
    def alignment(self, value):
        pPr = self.get_or_add_pPr()
        pPr.jc_val = value

    def clear_content(self):
        """
        Remove all child elements, except the ``<w:pPr>`` element if present.
        """
        for child in self[:]:
            if child.tag == qn('w:pPr'):
                continue
            self.remove(child)

    def set_sectPr(self, sectPr):
        """
        Unconditionally replace or add *sectPr* as a grandchild in the
        correct sequence.
        """
        pPr = self.get_or_add_pPr()
        pPr._remove_sectPr()
        pPr._insert_sectPr(sectPr)

    @property
    def style(self):
        """
        String contained in w:val attribute of ./w:pPr/w:pStyle grandchild,
        or |None| if not present.
        """
        pPr = self.pPr
        if pPr is None:
            return None
        return pPr.style

    @style.setter
    def style(self, style):
        pPr = self.get_or_add_pPr()
        pPr.style = style
